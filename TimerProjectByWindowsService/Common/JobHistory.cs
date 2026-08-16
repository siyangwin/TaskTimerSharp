using System;
using System.Globalization;
using System.IO;

namespace TimerProjectByWindowsService.Common
{
    /// <summary>
    /// 执行历史持久化：每次执行（成功/失败）追加写入 Log/History/任务名/yyyyMM.csv，
    /// 服务重启后可据此恢复"今天已执行过"的周期任务状态，避免当天重复执行。
    /// CSV格式：时间,触发方式,耗时秒,状态,结果摘要
    /// </summary>
    public static class JobHistory
    {
        private static readonly object WriteSync = new object();

        /// <summary>记录一次执行结果；写入失败不影响主流程</summary>
        public static void Record(string fun, string trigger, double seconds, string status, string summary)
        {
            try
            {
                string dir = Path.Combine(RootDir(), fun);
                Directory.CreateDirectory(dir);
                string file = Path.Combine(dir, DateTime.Now.ToString("yyyyMM") + ".csv");

                string line = string.Join(",",
                    Sanitize(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                    Sanitize(trigger),
                    seconds.ToString("0.0", CultureInfo.InvariantCulture),
                    Sanitize(status),
                    Sanitize(LocalFile.Truncate(summary ?? "", 500)));

                lock (WriteSync)
                {
                    File.AppendAllText(file, line + Environment.NewLine);
                }
            }
            catch
            {
                //历史写入失败不影响主流程
            }
        }

        /// <summary>今天是否已有成功执行记录（用于服务重启后恢复周期任务状态）</summary>
        public static bool HasSuccessToday(string fun)
        {
            try
            {
                string file = Path.Combine(RootDir(), fun, DateTime.Now.ToString("yyyyMM") + ".csv");
                if (!File.Exists(file)) return false;

                string today = DateTime.Now.ToString("yyyy-MM-dd");
                foreach (string line in File.ReadLines(file))
                {
                    if (line.StartsWith(today, StringComparison.Ordinal) && line.Contains(",成功,"))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private static string RootDir()
        {
            return Path.Combine(new LocalFile().LocalLog(), "History");
        }

        //字段清洗：CSV不含引号/逗号/换行，读取端可以简单按逗号分割
        private static string Sanitize(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Replace(',', '，').Replace('"', '\'').Replace('\r', ' ').Replace('\n', ' ');
        }
    }
}
