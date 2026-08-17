using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using System.Threading;

namespace TimerProjectByWindowsService.Common
{
    /// <summary>
    /// 专用日志写入器：有界队列 + 单写入线程 + 重复消息节流 + 单文件大小轮转 + 关机排空。
    /// 目录结构与行格式保持与旧版 WiterLog 完全一致：Log/&lt;任务&gt;/&lt;yyyyMM&gt;/&lt;yyyyMMdd&gt;.log，
    /// 行格式 [yyyy-MM-dd HH:mm:ss]       内容
    /// </summary>
    public static class LogWriter
    {
        private class LogEntry
        {
            public string FilePath;
            public string Line;
        }

        private class ThrottleState
        {
            public DateTime LastWriteAt;
            public int SuppressedCount;
        }

        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        //可选配置项(appSettings)，缺省值即可直接使用
        private static readonly int QueueSize = ReadIntSetting("LogQueueSize", 20000);
        private static readonly int ThrottleMinutes = ReadIntSetting("LogThrottleMinutes", 30);
        private static readonly long MaxFileBytes = ReadIntSetting("LogMaxMB", 50) * 1024L * 1024L;

        private static readonly BlockingCollection<LogEntry> Queue = new BlockingCollection<LogEntry>(QueueSize);
        private static readonly Thread WriterThread;

        private static readonly object ThrottleSync = new object();
        private static readonly Dictionary<string, ThrottleState> Throttle = new Dictionary<string, ThrottleState>();

        private static readonly object FailureSync = new object();
        private static int _consecutiveFailures;
        private static DateTime _lastFailureReport = DateTime.MinValue;

        private static readonly object DropSync = new object();
        private static long _droppedCount;
        private static DateTime _lastDropReport = DateTime.MinValue;

        /// <summary>由宿主(服务)注入的外部报告通道(写事件查看器)；写入失败/队列丢弃时用它留痕，避免彻底静默</summary>
        public static Action<string> ExternalReporter;

        static LogWriter()
        {
            WriterThread = new Thread(WriterLoop) { IsBackground = true, Name = "LogWriter" };
            WriterThread.Start();
        }

        /// <summary>入队一条日志。路径计算与节流在调用线程完成，入队本身不阻塞调用方</summary>
        public static void Enqueue(string fun, string info)
        {
            try
            {
                DateTime now = DateTime.Now;

                string suffix;
                if (!TryPassThrottle(fun, info, now, out suffix)) return;

                string basePath = new LocalFile().LocalLog();
                string dir = Path.Combine(Path.Combine(basePath, fun), now.ToString("yyyyMM"));
                string filePath = Path.Combine(dir, now.ToString("yyyyMMdd") + ".log");

                string line = "[" + now.ToString("yyyy-MM-dd HH:mm:ss") + "]       " + info + suffix + "\r\n";

                if (!Queue.TryAdd(new LogEntry { FilePath = filePath, Line = line }))
                {
                    NoteDropped(now);
                }
            }
            catch (Exception ex)
            {
                //日志入队失败不递归写日志，只尝试外部报告
                Report("LogWriter入队出现问题：" + ex.Message);
            }
        }

        /// <summary>关机排空：停止接收新日志，等待写入线程把队列写完(最多等timeout)</summary>
        public static void Shutdown(TimeSpan timeout)
        {
            try
            {
                Queue.CompleteAdding();
                WriterThread.Join(timeout);
            }
            catch
            {
                //关机路径不再抛异常
            }
        }

        private static void WriterLoop()
        {
            foreach (LogEntry entry in Queue.GetConsumingEnumerable())
            {
                try
                {
                    WriteEntry(entry);
                    lock (FailureSync) _consecutiveFailures = 0;
                }
                catch (Exception ex)
                {
                    OnWriteFailed(ex);
                }
            }
        }

        private static void WriteEntry(LogEntry entry)
        {
            string dir = Path.GetDirectoryName(entry.FilePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

            FileInfo fi = new FileInfo(entry.FilePath);
            if (fi.Exists && fi.Length > MaxFileBytes)
            {
                //轮转：改名为 <日期>.1.log(已存在则覆盖)，保留最近一个旧文件
                string rotated = Path.Combine(dir, Path.GetFileNameWithoutExtension(entry.FilePath) + ".1.log");
                if (File.Exists(rotated)) File.Delete(rotated);
                File.Move(entry.FilePath, rotated);
            }

            File.AppendAllText(entry.FilePath, entry.Line, Utf8NoBom);
        }

        /// <summary>重复消息节流：同任务同消息在窗口内只写第一条；窗口过后再次出现时带上被省略的次数</summary>
        private static bool TryPassThrottle(string fun, string info, DateTime now, out string suffix)
        {
            suffix = "";
            string key = (fun ?? "") + "|" + (info ?? "");
            lock (ThrottleSync)
            {
                if (Throttle.Count > 1000)
                {
                    List<string> expired = new List<string>();
                    foreach (var kv in Throttle)
                    {
                        if ((now - kv.Value.LastWriteAt).TotalMinutes >= ThrottleMinutes) expired.Add(kv.Key);
                    }
                    foreach (string k in expired) Throttle.Remove(k);
                }

                ThrottleState state;
                if (Throttle.TryGetValue(key, out state))
                {
                    if ((now - state.LastWriteAt).TotalMinutes < ThrottleMinutes)
                    {
                        state.SuppressedCount++;
                        return false;
                    }
                    if (state.SuppressedCount > 0)
                    {
                        suffix = "（过去" + ThrottleMinutes + "分钟内重复" + state.SuppressedCount + "次，已省略）";
                    }
                    state.LastWriteAt = now;
                    state.SuppressedCount = 0;
                    return true;
                }

                Throttle[key] = new ThrottleState { LastWriteAt = now };
                return true;
            }
        }

        private static void OnWriteFailed(Exception ex)
        {
            lock (FailureSync)
            {
                _consecutiveFailures++;
                DateTime now = DateTime.Now;
                if ((now - _lastFailureReport).TotalSeconds >= 60)
                {
                    _lastFailureReport = now;
                    Report("日志写入失败(最近连续" + _consecutiveFailures + "次)：" + ex.Message);
                }
            }
        }

        private static void NoteDropped(DateTime now)
        {
            Interlocked.Increment(ref _droppedCount);
            lock (DropSync)
            {
                if ((now - _lastDropReport).TotalMinutes >= 10)
                {
                    _lastDropReport = now;
                    long count = Interlocked.Exchange(ref _droppedCount, 0);
                    Report("日志队列已满，最近丢弃" + count + "条日志(写入速度跟不上或写入失败堆积)");
                }
            }
        }

        private static void Report(string message)
        {
            Action<string> reporter = ExternalReporter;
            if (reporter == null) return;
            try
            {
                reporter(message);
            }
            catch
            {
                //报告通道自身失败时彻底静默，避免递归
            }
        }

        private static int ReadIntSetting(string key, int defaultValue)
        {
            int value;
            if (int.TryParse(ConfigurationManager.AppSettings[key], out value) && value > 0) return value;
            return defaultValue;
        }
    }
}
