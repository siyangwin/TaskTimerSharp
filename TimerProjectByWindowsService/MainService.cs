using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TimerProjectByWindowsService.Common;

namespace TimerProjectByWindowsService
{
    public partial class MainService : ServiceBase
    {
        //任务类型
        private enum JobKind
        {
            Api,                //Type=0 调用单个API
            Sequence,           //Type=1 顺序执行(API+存储过程)
            StoredProcedure     //Type=2 独立执行存储过程
        }

        //主定时器
        private System.Timers.Timer TimerProjectJob;

        //读取公共配置
        public static readonly string ProgramName = ConfigurationManager.AppSettings["ProgramName"]; //程序名称,表明当前是哪只定时程序
        public static readonly string SendTo = ConfigurationManager.AppSettings["SendTo"];           //ACTION通知会加入以下邮件地址和程序配置地址一起发送
        public static readonly string InfoMessage = ConfigurationManager.AppSettings["InfoMessage"]; //Info通知,是否加入SendTo配置地址一起发送 true:发送 false:不发送

        //HTTP请求默认超时(毫秒)。可在Setup.xml用HttpTimeout(单位:秒)覆盖
        private const int DefaultHttpTimeoutMs = 10 * 60 * 1000;

        //实例化文件类
        private readonly LocalFile localFile = new LocalFile();
        private readonly SendMail Send = new SendMail();

        //job子定时器注册表(停止服务/移除任务时需要找到并停掉它们)
        private class JobState
        {
            public System.Timers.Timer Timer;
            public int TickRunning; //子定时器重入守卫:上一轮调度(含同步发邮件)未结束时跳过本tick
        }

        private readonly Dictionary<string, JobState> _jobTimers = new Dictionary<string, JobState>();
        private readonly object _jobTimersLock = new object();

        //主循环重入标记
        private int _mainTickRunning;

        //最近一次"每日重置"对应的日期
        private DateTime _lastResetDate = DateTime.Today;

        public MainService()
        {
            InitializeComponent();
        }

        #region 服务启动/停止

        //服务启动执行代码
        protected override void OnStart(string[] args)
        {
            try
            {
                //注入日志写入器的外部报告通道(写事件查看器)
                LogWriter.ExternalReporter = TryWriteEventLog;

                string ver = "未知";
                try { ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(); } catch { }
                TryWriteEventLog("我的服务启动 v" + ver);
                WiterLog("System", "我的服务启动 v" + ver);

                //初始化全局状态
                PassValue.ClearRunningJobs();
                PassValue.UnlockAllJobs();
                PassValue.ClearExecutedToday();
                _lastResetDate = DateTime.Today;

                //初始化主定时器,每秒扫描一次系统配置
                TimerProjectJob = new System.Timers.Timer();
                TimerProjectJob.Interval = 1000;
                TimerProjectJob.Elapsed += TimerProjectJob_Elapsed;
                TimerProjectJob.AutoReset = true;
                TimerProjectJob.Enabled = true;
            }
            catch (Exception ex)
            {
                TryWriteEventLog("服务启动失败：" + ex);
                SendActionEmail(null, "System", "SystemOnStart", ex.Message, ex.ToString());
            }
        }

        //服务停止执行代码
        protected override void OnStop()
        {
            try
            {
                //先停主定时器,不再派发新任务
                if (TimerProjectJob != null)
                {
                    TimerProjectJob.Enabled = false;
                    TimerProjectJob.Dispose();
                    TimerProjectJob = null;
                }

                //停掉所有job子定时器
                List<JobState> states;
                lock (_jobTimersLock)
                {
                    states = new List<JobState>(_jobTimers.Values);
                    _jobTimers.Clear();
                }
                foreach (var state in states)
                {
                    if (state.Timer != null)
                    {
                        state.Timer.Enabled = false;
                        state.Timer.Dispose();
                    }
                }
                PassValue.ClearRunningJobs();

                //等待在途任务结束,最多12秒(留余量给SCM的20秒强杀时限+日志排空)
                DateTime deadline = DateTime.Now.AddSeconds(12);
                while (PassValue.HasLockedJobs() && DateTime.Now < deadline)
                {
                    Thread.Sleep(500);
                }
                PassValue.UnlockAllJobs();

                TryWriteEventLog("我的服务停止");
                WiterLog("System", "我的服务停止");
                //等待日志排空(最多5秒)后再退出,保证最后一条日志落盘
                LogWriter.Shutdown(TimeSpan.FromSeconds(5));
            }
            catch (Exception ex)
            {
                TryWriteEventLog("停止服务出现问题：" + ex);
                WiterLog("System", "我正在执行停止运行,出现问题：" + ex);
                SendActionEmail(null, "System", "SystemOnStop", ex.Message, ex.ToString());
            }
        }

        #endregion

        #region 主循环：扫描系统配置,启动/停止Job

        private void TimerProjectJob_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //防止上一轮还没跑完时重入
            if (Interlocked.CompareExchange(ref _mainTickRunning, 1, 0) != 0)
            {
                return;
            }

            try
            {
                //日期变化时清空"今天已执行"记录(不再依赖恰好命中23:59:59那一秒)
                if (DateTime.Today != _lastResetDate)
                {
                    PassValue.ClearExecutedToday();
                    _lastResetDate = DateTime.Today;
                    WiterLog("System", "日期变更,已清空周期任务今日执行记录");
                }

                //获取配置文件根目录,拼接System配置文件路径
                string basePath = localFile.LocalConfig();
                string configPath = Path.Combine(basePath, "System");
                if (!Directory.Exists(configPath))
                {
                    Directory.CreateDirectory(configPath);
                }

                DataTable dt = localFile.GetXmlInfo(configPath, "TimeProjectJob.xml");
                if (dt == null || dt.Rows.Count == 0)
                {
                    //没有配置数据,停止所有job
                    foreach (string running in PassValue.GetRunningJobs())
                    {
                        StopJob(running, true);
                    }
                    return;
                }

                List<string> activeJobs = new List<string>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string nameCN = GetRowField(dt, i, "NameCN");
                    string nameEN = GetRowField(dt, i, "NameEN");
                    string status = GetRowField(dt, i, "Status");
                    string type = GetRowField(dt, i, "Type");

                    if (nameEN == "" || !status.Equals("true", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    JobKind kind;
                    if (type == "0") kind = JobKind.Api;
                    else if (type == "1") kind = JobKind.Sequence;
                    else if (type == "2") kind = JobKind.StoredProcedure;
                    else
                    {
                        WiterLog("System", "TimeProjectJob.xml中[" + nameCN + "(" + nameEN + ")]配置Type错误(支持0/1/2),无法执行");
                        SendActionEmail(dt, "System", "System", "System配置错误", "TimeProjectJob.xml中[" + nameCN + "(" + nameEN + ")]配置Type错误(支持0/1/2),无法执行");
                        continue;
                    }

                    activeJobs.Add(nameEN);
                    StartJob(nameEN, nameCN, kind);
                }

                //停止已不在配置中(或被改为停止)的job
                foreach (string running in PassValue.GetRunningJobs())
                {
                    if (!activeJobs.Contains(running))
                    {
                        StopJob(running, true);
                    }
                }
            }
            catch (Exception ex)
            {
                TryWriteEventLog("主循环出现问题：" + ex.Message);
                WiterLog("System", "我正在执行System_Time,出现问题：" + ex);
                SendActionEmail(null, "System", "System", ex.Message, ex.ToString());
            }
            finally
            {
                Interlocked.Exchange(ref _mainTickRunning, 0);
            }
        }

        //为一个任务启动独立子定时器(已启动则忽略)
        private void StartJob(string fun, string nameCN, JobKind kind)
        {
            lock (_jobTimersLock)
            {
                if (_jobTimers.ContainsKey(fun)) return;
            }

            //周期任务:根据执行历史恢复"今日已执行"状态,防止服务重启后当天重复执行
            RestoreExecutedTodayFromHistory(fun);

            var state = new JobState();
            state.Timer = new System.Timers.Timer(1000) { AutoReset = true };
            state.Timer.Elapsed += (o, e) => JobTimer_Elapsed(fun, kind, state);

            lock (_jobTimersLock)
            {
                if (_jobTimers.ContainsKey(fun))
                {
                    state.Timer.Dispose();
                    return;
                }
                _jobTimers[fun] = state;
            }

            PassValue.AddRunningJob(fun);
            state.Timer.Enabled = true;
            WiterLog("System", "开始运行" + KindName(kind) + "：" + nameCN + "(" + fun + ")");
        }

        //停止一个任务的子定时器并清理其运行状态
        private void StopJob(string fun, bool log)
        {
            JobState state = null;
            lock (_jobTimersLock)
            {
                if (_jobTimers.TryGetValue(fun, out state))
                {
                    _jobTimers.Remove(fun);
                }
            }

            PassValue.RemoveRunningJob(fun);

            if (state != null && state.Timer != null)
            {
                state.Timer.Enabled = false;
                state.Timer.Dispose();
            }

            PassValue.RemoveLastRunTime(fun);
            PassValue.UnmarkExecutedToday(fun);
            //不再在这里解锁:如果任务仍在执行中,提前解锁会导致两次执行并发运行;
            //执行器的finally保证解锁,OnStop末尾的UnlockAllJobs是兜底

            if (log)
            {
                WiterLog("System", "停止运行" + fun + "!");
            }
        }

        //服务重启后,根据执行历史判断周期任务今天是否已成功执行过
        private void RestoreExecutedTodayFromHistory(string fun)
        {
            try
            {
                string configPath = Path.Combine(localFile.LocalConfig(), fun);
                DataTable dt = localFile.GetXmlInfo(configPath, fun + "_Setup.xml");
                if (dt == null || dt.Rows.Count == 0) return;

                string mode = GetField(dt, "ExecutionStatus");
                if (mode == "1" && JobHistory.HasSuccessToday(fun))
                {
                    PassValue.MarkExecutedToday(fun);
                    WiterLog(fun, "根据执行历史,今天已成功执行过该周期任务,今日不再重复执行");
                }
            }
            catch (Exception ex)
            {
                WiterLog(fun, "恢复今日执行状态失败：" + ex.Message);
            }
        }

        private static string KindName(JobKind kind)
        {
            switch (kind)
            {
                case JobKind.Api: return "调用API";
                case JobKind.Sequence: return "顺序执行";
                case JobKind.StoredProcedure: return "执行存储过程";
                default: return kind.ToString();
            }
        }

        #endregion

        #region Job调度（三种任务类型共用）

        //每个job子定时器的调度入口:读取Setup.xml,判断是否到了执行时间
        private void JobTimer_Elapsed(string fun, JobKind kind, JobState state)
        {
            //上一轮调度(含同步发邮件)还没跑完时跳过本tick,防止线程池堆积
            if (Interlocked.CompareExchange(ref state.TickRunning, 1, 0) != 0)
            {
                return;
            }

            try
            {
                //任务已被主循环移除,自行停止
                if (!PassValue.ContainsRunningJob(fun))
                {
                    WiterLog(fun, "停止运行" + KindName(kind) + fun + "-Time!");
                    StopJob(fun, false);
                    return;
                }

                string configPath = Path.Combine(localFile.LocalConfig(), fun);
                if (!Directory.Exists(configPath))
                {
                    Directory.CreateDirectory(configPath);
                }

                DataTable dt = localFile.GetXmlInfo(configPath, fun + "_Setup.xml");
                if (dt == null || dt.Rows.Count == 0)
                {
                    return;
                }

                DateTime now = DateTime.Now;

                //开始时间(可选)
                string startTime = GetField(dt, "StartTime");
                if (!string.IsNullOrWhiteSpace(startTime))
                {
                    DateTime parsedStart;
                    if (!TryParseDateTime(startTime, out parsedStart))
                    {
                        ConfigAlert(dt, fun, fun + "-Time", fun + "_Setup.xml中StartTime格式错误(应为yyyy-MM-dd HH:mm),当前值：" + startTime);
                        return;
                    }
                    if (now < parsedStart) return; //未到开始时间,静默等待
                }

                //结束时间(可选)
                string endTime = GetField(dt, "EndTime");
                if (!string.IsNullOrWhiteSpace(endTime))
                {
                    DateTime parsedEnd;
                    if (!TryParseDateTime(endTime, out parsedEnd))
                    {
                        ConfigAlert(dt, fun, fun + "-Time", fun + "_Setup.xml中EndTime格式错误(应为yyyy-MM-dd HH:mm),当前值：" + endTime);
                        return;
                    }
                    if (now > parsedEnd) return; //已过结束时间,静默等待
                }

                string mode = GetField(dt, "ExecutionStatus");
                if (mode == "0")
                {
                    ScheduleByInterval(fun, kind, dt, now);
                }
                else if (mode == "1")
                {
                    ScheduleByCycle(fun, kind, dt, now);
                }
                else
                {
                    ConfigAlert(dt, fun, fun + "-Time", fun + "_Setup.xml中ExecutionStatus配置错误,无法执行。请按照说明调整(执行类型 0:按时间间隔 1:按周期)。");
                }
            }
            catch (Exception ex)
            {
                TryWriteEventLog(fun + " 调度出现问题：" + ex.Message);
                WiterLog(fun, "调度出现问题：" + ex);
                SendActionEmail(null, fun, fun + "-Time", ex.Message, ex.ToString());
            }
            finally
            {
                Interlocked.Exchange(ref state.TickRunning, 0);
            }
        }

        //按时间间隔调度
        private void ScheduleByInterval(string fun, JobKind kind, DataTable dt, DateTime now)
        {
            string intervals = GetField(dt, "IntervalsTime");
            double intervalSeconds;
            if (string.IsNullOrWhiteSpace(intervals)
                || !double.TryParse(intervals, NumberStyles.Any, CultureInfo.InvariantCulture, out intervalSeconds)
                || intervalSeconds <= 0)
            {
                ConfigAlert(dt, fun, fun + "-Time", fun + "_Setup.xml中IntervalsTime配置错误,无法执行。请按照说明调整(时间间隔 单位:秒,必须为正数)。");
                return;
            }

            DateTime lastRun;
            if (!PassValue.TryGetLastRunTime(fun, out lastRun))
            {
                QueueExecution(fun, kind, dt, "时间间隔,第一次执行");
                return;
            }

            //上次执行完成时间+间隔 已到,触发执行(执行中会被执行锁挡住,不会并发)
            if (lastRun.AddSeconds(intervalSeconds) <= now)
            {
                QueueExecution(fun, kind, dt, "时间间隔执行");
            }
        }

        //按周期调度(EveryDay/EveryWeek)
        private void ScheduleByCycle(string fun, JobKind kind, DataTable dt, DateTime now)
        {
            string cycleType = GetField(dt, "CycltType");

            if (cycleType.Equals("everyweek", StringComparison.OrdinalIgnoreCase))
            {
                string dayOfWeekStr = GetField(dt, "DayOfWeek");
                DayOfWeek dayOfWeek;
                if (!Enum.TryParse(dayOfWeekStr, true, out dayOfWeek))
                {
                    ConfigAlert(dt, fun, fun + "-Time", fun + "_Setup.xml中DayOfWeek配置错误(应为Monday~Sunday的英文),当前值：" + dayOfWeekStr);
                    return;
                }
                if (dayOfWeek != now.DayOfWeek) return;
            }
            else if (!cycleType.Equals("everyday", StringComparison.OrdinalIgnoreCase))
            {
                ConfigAlert(dt, fun, fun + "-Time", fun + "_Setup.xml中CycltType配置错误,无法执行。请按照说明调整(周期选择:EveryDay EveryWeek)。");
                return;
            }

            string specificTime = GetField(dt, "SpecificTime");
            TimeSpan timeOfDay;
            if (!TryParseTimeOfDay(specificTime, out timeOfDay))
            {
                ConfigAlert(dt, fun, fun + "-Time", fun + "_Setup.xml中SpecificTime配置错误(应为HH:mm,24小时制),当前值：" + specificTime);
                return;
            }

            //到达执行分钟,原子地抢占"今天只执行一次"
            if (now.Hour == timeOfDay.Hours && now.Minute == timeOfDay.Minutes)
            {
                if (PassValue.MarkExecutedToday(fun))
                {
                    QueueExecution(fun, kind, dt, "周期执行");
                }
            }
        }

        //把真正的执行工作放到线程池,调度线程不被阻塞;执行锁保证同一任务不会并发执行
        private void QueueExecution(string fun, JobKind kind, DataTable dt, string info)
        {
            Task.Run(() =>
            {
                try
                {
                    switch (kind)
                    {
                        case JobKind.Api:
                            ApiInfoLoadData(fun, dt, info);
                            break;
                        case JobKind.Sequence:
                            APISequenceLoadData(fun, dt, info);
                            break;
                        case JobKind.StoredProcedure:
                            SPInfoLoadData(fun, dt, info);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    WiterLog(fun, "执行出现未捕获异常：" + ex);
                    JobHistory.Record(fun, info, 0, "失败", ex.Message);
                    SendActionEmail(dt, fun, fun + "-LoadData", ex.Message, ex.ToString());
                    PassValue.UnlockJob(fun);
                }
            });
        }

        #endregion

        #region 执行器

        /// <summary>Type=0 调用单个API</summary>
        private void ApiInfoLoadData(string fun, DataTable dt, string info)
        {
            if (!PassValue.LockJob(fun))
            {
                WiterLog(fun, "程序还未处理完成，锁定执行");
                return;
            }

            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                string apiUrl = GetField(dt, "ApiUrl");
                if (string.IsNullOrWhiteSpace(apiUrl))
                {
                    ConfigAlertFail(dt, fun, fun + "-ApiInfo", info, sw, fun + "_Setup.xml中ApiUrl配置错误,无法执行。请按照说明调整(Api 请求地址)。");
                    return;
                }

                string url;
                if (!TryBuildSignedUrl(dt, apiUrl, out url))
                {
                    ConfigAlertFail(dt, fun, fun + "-ApiInfo", info, sw, fun + "_Setup.xml中AuthenticationKey配置错误,无法执行。请按照说明调整(请求安全验证之密钥)。");
                    return;
                }

                int retryCount, retryInterval;
                GetRetryConfig(dt, out retryCount, out retryInterval);
                int timeoutMs = GetHttpTimeoutMs(dt);

                WiterLog(fun, info + ",程序开始执行!");
                string result = ExecuteWithRetry(fun, () => HttpRequest.HttpGet(url, "", timeoutMs), retryCount, retryInterval);
                sw.Stop();

                WiterLog(fun, "返回数据：" + LocalFile.Truncate(result, 4000));
                JobHistory.Record(fun, info, sw.Elapsed.TotalSeconds, "成功", result);
                SendInfoEmail(dt, fun, fun + "-LoadData", "執行成功", LocalFile.Truncate(result, 2000));
            }
            catch (Exception ex)
            {
                sw.Stop();
                TryWriteEventLog(fun + " LoadData出现问题：" + ex.Message);
                WiterLog(fun, "执行出现问题：" + ex);
                JobHistory.Record(fun, info, sw.Elapsed.TotalSeconds, "失败", ex.Message);
                SendActionEmail(dt, fun, fun + "-LoadData", ex.Message, ex.ToString());
            }
            finally
            {
                //执行完成(无论成败)才解锁并刷新计时基准,避免失败后立即重试风暴
                PassValue.SetLastRunTime(fun, DateTime.Now);
                PassValue.UnlockJob(fun);
            }
        }

        /// <summary>Type=1 顺序执行(API+存储过程),任一步失败立即中断后续步骤</summary>
        private void APISequenceLoadData(string fun, DataTable dt, string info)
        {
            if (!PassValue.LockJob(fun))
            {
                WiterLog(fun, "程序还未处理完成，锁定执行");
                return;
            }

            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                string configPath = Path.Combine(localFile.LocalConfig(), fun);
                DataTable sequenceDT = localFile.GetXmlInfo(configPath, fun + "_Sequence.xml");
                if (sequenceDT == null || sequenceDT.Rows.Count == 0)
                {
                    ConfigAlertFail(dt, fun, fun + "-APISequence", info, sw, fun + "_Sequence.xml不存在或没有配置步骤,无法执行。");
                    return;
                }

                int retryCount, retryInterval;
                GetRetryConfig(dt, out retryCount, out retryInterval);
                int timeoutMs = GetHttpTimeoutMs(dt);

                WiterLog(fun, info + ",程序开始执行!");
                StringBuilder aliResult = new StringBuilder();

                for (int i = 0; i < sequenceDT.Rows.Count; i++)
                {
                    string stepType = sequenceDT.Rows[i][0].ToString().Trim();
                    string stepInfo = sequenceDT.Rows[i][1].ToString().Trim();
                    WiterLog(fun, "执行第" + (i + 1) + "个" + stepType + "：" + stepInfo);

                    string res;
                    if (stepType == "存储过程")
                    {
                        string connStr = GetField(dt, "ConnStr");
                        if (string.IsNullOrWhiteSpace(connStr))
                        {
                            ConfigAlertFail(dt, fun, fun + "-APISequence", info, sw, "第" + (i + 1) + "步为存储过程,但" + fun + "_Setup.xml中ConnStr为空,无法执行。请按照说明调整(数据库链接字符串)。");
                            return;
                        }

                        string spName = stepInfo;
                        DataTable dtSql = ExecuteWithRetry(fun, () => SQLHelper.ExecuteDataTable(connStr, CommandType.StoredProcedure, spName, null), retryCount, retryInterval);
                        res = (dtSql != null && dtSql.Rows.Count > 0) ? dtSql.Rows[0][0].ToString() : "(无返回数据)";
                    }
                    else if (stepType == "API地址")
                    {
                        if (string.IsNullOrWhiteSpace(stepInfo))
                        {
                            ConfigAlertFail(dt, fun, fun + "-APISequence", info, sw, "第" + (i + 1) + "步API地址为空,无法执行。请调整" + fun + "_Sequence.xml。");
                            return;
                        }

                        string url;
                        if (!TryBuildSignedUrl(dt, stepInfo, out url))
                        {
                            ConfigAlertFail(dt, fun, fun + "-APISequence", info, sw, "第" + (i + 1) + "步启用了安全验证,但" + fun + "_Setup.xml中AuthenticationKey为空,无法执行。");
                            return;
                        }

                        res = ExecuteWithRetry(fun, () => HttpRequest.HttpGet(url, "", timeoutMs), retryCount, retryInterval);
                    }
                    else
                    {
                        ConfigAlertFail(dt, fun, fun + "-APISequence", info, sw, "第" + (i + 1) + "步类型[" + stepType + "]不支持(仅支持:API地址/存储过程),顺序执行已中断。");
                        return;
                    }

                    string stepResult = "序号：" + (i + 1) + " " + stepType + ":" + stepInfo + "；返回数据：" + LocalFile.Truncate(res, 500);
                    aliResult.AppendLine(stepResult);
                    WiterLog(fun, stepResult);
                }

                sw.Stop();
                string resultText = aliResult.ToString();
                JobHistory.Record(fun, info, sw.Elapsed.TotalSeconds, "成功", resultText);
                SendInfoEmail(dt, fun, fun + "-LoadData", "執行成功", LocalFile.Truncate(resultText, 2000));
            }
            catch (Exception ex)
            {
                sw.Stop();
                WiterLog(fun, "发生错误：" + ex);
                JobHistory.Record(fun, info, sw.Elapsed.TotalSeconds, "失败", ex.Message);
                SendActionEmail(dt, fun, fun + "-LoadData", ex.Message, ex.ToString());
            }
            finally
            {
                PassValue.SetLastRunTime(fun, DateTime.Now);
                PassValue.UnlockJob(fun);
            }
        }

        /// <summary>Type=2 独立执行存储过程</summary>
        private void SPInfoLoadData(string fun, DataTable dt, string info)
        {
            if (!PassValue.LockJob(fun))
            {
                WiterLog(fun, "程序还未处理完成，锁定执行");
                return;
            }

            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                string connStr = GetField(dt, "ConnStr");
                if (string.IsNullOrWhiteSpace(connStr))
                {
                    ConfigAlertFail(dt, fun, fun + "-SP", info, sw, fun + "_Setup.xml中ConnStr配置错误,无法执行。请按照说明调整(数据库链接字符串)。");
                    return;
                }

                string spName = GetField(dt, "StoredProcedure");
                if (string.IsNullOrWhiteSpace(spName))
                {
                    ConfigAlertFail(dt, fun, fun + "-SP", info, sw, fun + "_Setup.xml中StoredProcedure未配置,无法执行。请填写存储过程名称。");
                    return;
                }

                int retryCount, retryInterval;
                GetRetryConfig(dt, out retryCount, out retryInterval);

                WiterLog(fun, info + ",程序开始执行存储过程：" + spName);
                DataTable dtSql = ExecuteWithRetry(fun, () => SQLHelper.ExecuteDataTable(connStr, CommandType.StoredProcedure, spName, null), retryCount, retryInterval);
                sw.Stop();

                string res = (dtSql != null && dtSql.Rows.Count > 0) ? dtSql.Rows[0][0].ToString() : "(无返回数据)";
                WiterLog(fun, "返回数据：" + LocalFile.Truncate(res, 4000));
                JobHistory.Record(fun, info, sw.Elapsed.TotalSeconds, "成功", res);
                SendInfoEmail(dt, fun, fun + "-LoadData", "執行成功", LocalFile.Truncate(res, 2000));
            }
            catch (Exception ex)
            {
                sw.Stop();
                WiterLog(fun, "发生错误：" + ex);
                JobHistory.Record(fun, info, sw.Elapsed.TotalSeconds, "失败", ex.Message);
                SendActionEmail(dt, fun, fun + "-LoadData", ex.Message, ex.ToString());
            }
            finally
            {
                PassValue.SetLastRunTime(fun, DateTime.Now);
                PassValue.UnlockJob(fun);
            }
        }

        #endregion

        #region 重试与配置辅助

        /// <summary>带重试的执行;全部重试失败才向上抛异常</summary>
        private T ExecuteWithRetry<T>(string fun, Func<T> action, int retryCount, int retryIntervalSeconds)
        {
            int attempt = 0;
            while (true)
            {
                try
                {
                    return action();
                }
                catch (Exception ex)
                {
                    attempt++;
                    if (attempt > retryCount) throw;
                    WiterLog(fun, "第" + attempt + "次执行失败：" + ex.Message + "，" + retryIntervalSeconds + "秒后重试");
                    Thread.Sleep(retryIntervalSeconds * 1000);
                }
            }
        }

        //读取可选的重试配置 RetryCount(默认0=不重试) / RetryInterval(秒,默认60)
        private static void GetRetryConfig(DataTable dt, out int retryCount, out int retryIntervalSeconds)
        {
            retryCount = 0;
            retryIntervalSeconds = 60;
            int.TryParse(GetField(dt, "RetryCount"), NumberStyles.Integer, CultureInfo.InvariantCulture, out retryCount);
            int.TryParse(GetField(dt, "RetryInterval"), NumberStyles.Integer, CultureInfo.InvariantCulture, out retryIntervalSeconds);
            if (retryCount < 0) retryCount = 0;
            if (retryCount > 10) retryCount = 10;
            if (retryIntervalSeconds < 1) retryIntervalSeconds = 60;
        }

        //读取可选的HTTP超时配置 HttpTimeout(秒),缺省10分钟
        private static int GetHttpTimeoutMs(DataTable dt)
        {
            int seconds;
            if (int.TryParse(GetField(dt, "HttpTimeout"), NumberStyles.Integer, CultureInfo.InvariantCulture, out seconds) && seconds > 0)
            {
                return seconds * 1000;
            }
            return DefaultHttpTimeoutMs;
        }

        /// <summary>按配置追加签名参数;启用验证但密钥为空时返回false</summary>
        private static bool TryBuildSignedUrl(DataTable dt, string apiUrl, out string url)
        {
            url = apiUrl;
            string verification = GetField(dt, "Verification");
            if (!verification.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;

            string key = GetField(dt, "AuthenticationKey");
            if (key == "") return false;

            string time = LocalFile.GetTimeStampByMilliseconds();
            string sep = apiUrl.Contains("?") ? "&" : "?";
            url = apiUrl + sep + "Timestamp=" + time + "&SyncKey=" + LocalFile.GetSha1(key + time);
            return true;
        }

        //配置错误:写日志+发告警邮件(告警邮件在SendActionEmail内统一节流)
        private void ConfigAlert(DataTable dt, string fun, string funDetails, string message)
        {
            WiterLog(fun, message);
            SendActionEmail(dt, fun, funDetails, "配置错误", message);
        }

        /// <summary>执行器路径的配置错误:在ConfigAlert基础上补写执行历史(调度层不调用此方法,避免每秒触发刷爆CSV)</summary>
        private void ConfigAlertFail(DataTable dt, string fun, string funDetails, string info, Stopwatch sw, string message)
        {
            ConfigAlert(dt, fun, funDetails, message);
            JobHistory.Record(fun, info, sw.Elapsed.TotalSeconds, "失败", message);
        }

        private static bool TryParseDateTime(string value, out DateTime result)
        {
            return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
        }

        private static bool TryParseTimeOfDay(string value, out TimeSpan result)
        {
            result = TimeSpan.Zero;
            if (string.IsNullOrWhiteSpace(value)) return false;

            DateTime parsed;
            if (DateTime.TryParseExact(value.Trim(), new[] { "HH:mm", "H:mm", "HH:mm:ss" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
            {
                result = parsed.TimeOfDay;
                return true;
            }
            return false;
        }

        //安全读取配置字段:列不存在或值为空时返回""而不是抛异常
        private static string GetField(DataTable dt, string name)
        {
            return GetRowField(dt, 0, name);
        }

        private static string GetRowField(DataTable dt, int rowIndex, string name)
        {
            if (dt == null || rowIndex >= dt.Rows.Count) return "";
            if (!dt.Columns.Contains(name)) return "";
            object value = dt.Rows[rowIndex][name];
            return value == null || value == DBNull.Value ? "" : value.ToString().Trim();
        }

        #endregion

        #region 写入日志

        /// <summary>写入日志:委托给专用写入器(有界队列+单线程写+节流+轮转+关机排空)</summary>
        public void WiterLog(string Fun, string Info)
        {
            LogWriter.Enqueue(Fun, Info);
        }

        private void TryWriteEventLog(string message)
        {
            try
            {
                EventLog.WriteEntry(message);
            }
            catch
            {
                //事件源未注册等情况下忽略
            }
        }

        #endregion

        #region 发送邮件

        /// <summary>
        /// 程序运行错误,发送紧急邮件 [ACTION] - XXXX
        /// 同类错误在节流窗口内只发一封,避免邮件风暴
        /// </summary>
        public void SendActionEmail(DataTable dt, string Fun, string FunDetails, string Message, string Details)
        {
            string throttleKey = (Fun ?? "") + "|" + (FunDetails ?? "") + "|" + (Message ?? "");
            if (!MailThrottle.Allow(throttleKey))
            {
                WiterLog(string.IsNullOrEmpty(Fun) ? "System" : Fun, "[ACTION]邮件已节流(" + MailThrottle.ThrottleMinutes + "分钟内同类错误只发一封)：" + FunDetails + " - " + Message);
                return;
            }

            string MailTo = "";
            if (dt != null && dt.Rows.Count > 0)
            {
                MailTo = GetField(dt, "MailTo");
            }

            //合并任务配置的收件人和全局SendTo
            if (!string.IsNullOrWhiteSpace(SendTo))
            {
                MailTo = string.IsNullOrWhiteSpace(MailTo) ? SendTo : MailTo + ";" + SendTo;
            }

            if (string.IsNullOrWhiteSpace(MailTo))
            {
                WiterLog(string.IsNullOrEmpty(Fun) ? "System" : Fun, "收件人为空,邮件无法发送!");
                WiterLog("Mail", FunDetails + "收件人为空,邮件无法发送!");
                return;
            }

            WiterLog(string.IsNullOrEmpty(Fun) ? "System" : Fun, "收件人为：" + MailTo.Trim());
            WiterLog("Mail", FunDetails + " 收件人为：" + MailTo.Trim());

            string Subject = " [ACTION]-[" + ProgramName + "]-[" + FunDetails + "]-[" + Message + "]";

            string Body = "<table>";
            Body += "<tr style=\"height:40px;\"><td style=\"width:80px;text-align:left;vertical-align:top;\">Dear All:</td><td style=\"width:80px;text-align:left;vertical-align:top;\"></td><td></td></tr>";
            Body += "<tr style=\"height:40px;\"><td style=\"text-align:left;vertical-align:top;\"></td><td style=\"text-align:left;vertical-align:top;\" colspan=\"2\"> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 程序出現異常情況，請緊急處理！！！</td></tr>";
            Body += "<tr style=\"height:40px;\"><td style=\"text-align:left;vertical-align:top;\"></td><td style=\"text-align:left;vertical-align:top;\" colspan=\"2\">異常情況描述:" + Message + " </td></tr>";
            Body += "<tr style=\"height:40px;\"><td style=\"text-align:left;vertical-align:top;\"></td><td style=\"text-align:left;vertical-align:top;\" colspan=\"2\">異常情況詳情:<br/>" + Details + "</td></tr></table>";

            string Res = Send.SendmailFile(MailTo.Trim(), new string[] { }, Subject, Body, MailPriority.High);
            WiterLog(string.IsNullOrEmpty(Fun) ? "System" : Fun, Res);
            WiterLog("Mail", FunDetails + "," + Res);
        }

        /// <summary>
        /// 程序运行成功,发送邮件通知 [INFO] - XXXX
        /// </summary>
        public void SendInfoEmail(DataTable dt, string Fun, string FunDetails, string Message, string Details)
        {
            string MailTo = "";
            if (dt != null && dt.Rows.Count > 0)
            {
                string sendMail = GetField(dt, "SendMail");
                if (!sendMail.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    //不需要發送郵件
                    return;
                }
                MailTo = GetField(dt, "MailTo");
            }

            //InfoMessage为true时,合并全局SendTo地址
            if (string.Equals(InfoMessage, "true", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(SendTo))
            {
                MailTo = string.IsNullOrWhiteSpace(MailTo) ? SendTo : MailTo + ";" + SendTo;
            }

            if (string.IsNullOrWhiteSpace(MailTo))
            {
                WiterLog(Fun, "收件人为空,邮件无法发送!");
                WiterLog("Mail", FunDetails + "收件人为空,邮件无法发送!");
                return;
            }

            WiterLog(Fun, "收件人为：" + MailTo.Trim());
            WiterLog("Mail", FunDetails + " 收件人为：" + MailTo.Trim());

            string Subject = " [INFO]-[" + ProgramName + "]-[" + FunDetails + "]-[" + Message + "]";

            string Body = "<table>";
            Body += "<tr style=\"height:40px;\"><td style=\"width:80px;text-align:left;vertical-align:top;\">Dear All:</td><td style=\"width:80px;text-align:left;vertical-align:top;\"></td><td></td></tr>";
            Body += "<tr style=\"height:40px;\"><td style=\"text-align:left;vertical-align:top;\"></td><td style=\"text-align:left;vertical-align:top;\" colspan=\"2\"> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 程序執行成功</td></tr>";
            Body += "<tr style=\"height:40px;\"><td style=\"text-align:left;vertical-align:top;\"></td><td style=\"text-align:left;vertical-align:top;\" colspan=\"2\">執行結果:" + Message + " </td></tr>";
            Body += "<tr style=\"height:40px;\"><td style=\"text-align:left;vertical-align:top;\"></td><td style=\"text-align:left;vertical-align:top;\" colspan=\"2\">執行結果詳情:<br/> " + Details + "</td></tr></table>";

            string Res = Send.SendmailFile(MailTo.Trim(), new string[] { }, Subject, Body, MailPriority.Normal);
            WiterLog(Fun, Res);
            WiterLog("Mail", FunDetails + "," + Res);
        }

        #endregion
    }
}
