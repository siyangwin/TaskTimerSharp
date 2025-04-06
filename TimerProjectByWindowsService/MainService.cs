using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        //创建一个系统定时器
        System.Timers.Timer TimerProjectJob;  //计时器

        //读取公共配置
        public static readonly string ProgramName = ConfigurationSettings.AppSettings["ProgramName"]; //程序名称,表明当前是哪只定时程序
        public static readonly string SendTo = ConfigurationSettings.AppSettings["SendTo"];           //ACTION通知会加入以下邮件地址和程序配置地址一起发送
        public static readonly string InfoMessage = ConfigurationSettings.AppSettings["InfoMessage"]; //Info通知,是否加入SendTo配置地址一起发送 true:发送 false:不发送

        //实例化文件类
        LocalFile File = new LocalFile();
        SendMail Send = new SendMail();
        public MainService()
        {
            InitializeComponent();
        }

        //判断执行状态,开始执行标记为1，执行完后标记为0
        public static int OnStartStatus=0;

        //服务启动执行代码
        protected override void OnStart(string[] args)
        {
            try
            {
                EventLog.WriteEntry("我的服务启动");//在系统事件查看器里的应用程序事件里来源的描述
                WiterLog("System", "我的服务启动" + Thread.CurrentThread.ManagedThreadId.ToString("00"));

                //初始化 全局变量
                PassValue.TimeProject = new List<string>();//存储目前所有正在运行的Job 系统级别

                PassValue.TimeProjectJob = new List<string>();//存储目前正在运行中的Job  避免重复运行同一方法

                //初始化
                TimerProjectJob = new System.Timers.Timer();
                //设定定时时间  最小单位是1秒
                TimerProjectJob.Interval = 1000;  //设置计时器事件间隔执行时间

                //设置定时执行方法  到达时间的时候执行事件； 
                TimerProjectJob.Elapsed += new System.Timers.ElapsedEventHandler(TimerProjectJob_Elapsed);

                TimerProjectJob.AutoReset = true;//设置是执行一次（false）还是一直执行(true)； 

                TimerProjectJob.Enabled = true; //是否执行System.Timers.Timer.Elapsed事件；
            }
            catch (Exception ex)
            {
                //执行出现问题,发邮件给管理员
                //[ACTION] -XXXX
                SendActionEmail(null, "System", "SystemOnStart", ex.Message,ex.ToString());
            }
        }

        //主程序开始运行 Main program
        private void TimerProjectJob_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                //每天晚上到23:59:59的时候清空 按周期执行的存储数据， 便于第二天继续执行
                if (DateTime.Now.ToString("HH:mm:ss")=="23:59:59")
                {
                    PassValue.ExecutionStatus = new List<string>();
                }

                if (OnStartStatus == 1)
                {
                    WiterLog("System1", "上一轮还没有执行完毕,等待。");
                    return;
                }

                OnStartStatus =1;

                //检查本地需要启动多少Job程序
                //获取配置文件根目录
                string basePath = File.LocalConfig();

                //拼接System应该存在的配置文件路径
                string ConfigPath = Path.Combine(basePath, "System");

                //检查目录是否存在,不存在则创建
                if (File.DirectoryIsExist(ConfigPath))
                {
                    //获取Xml配置的信息
                    DataTable dt = File.GetXmlInfo(ConfigPath, "TimeProjectJob.xml");

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        //记录所有的Job
                        List<string> TimeProjectJob = new List<string>();

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string NameCN = dt.Rows[i]["NameCN"].ToString().Trim();
                            string NameEN = dt.Rows[i]["NameEN"].ToString().Trim();
                            string Status = dt.Rows[i]["Status"].ToString().ToLower().Trim();
                            string Type = dt.Rows[i]["Type"].ToString().Trim();

                            //WiterLog("System", "NameCN：" + NameCN + ";NameEN:"+NameEN + ";Status:"+ Status+ ";Type:"+ Type);

                            if (NameEN != "" && Status == "true")
                            {
                                if (Type == "0")
                                {
                                    //将可以正常执行的Job写入List，用于去掉需要排除掉的Job
                                    TimeProjectJob.Add(NameEN);
                                    if (PassValue.TimeProject == null || !PassValue.TimeProject.Contains(NameEN))
                                    {
                                        //添加到列表上储存起来
                                        PassValue.TimeProject.Add(NameEN);
                                        WiterLog("System", "开始运行ApiInfo：" + NameCN + "(" + NameEN + ")" + Thread.CurrentThread.ManagedThreadId.ToString("00"));
                                        ApiInfo(NameEN);
                                        System.Threading.Thread.Sleep(1000);
                                    }
                                }
                                else if (Type == "1")
                                {
                                    //将可以正常执行的Job写入List，用于去掉需要排除掉的Job
                                    TimeProjectJob.Add(NameEN);
                                    if (PassValue.TimeProject == null || !PassValue.TimeProject.Contains(NameEN))
                                    {
                                        //添加到列表上储存起来
                                        PassValue.TimeProject.Add(NameEN);
                                        WiterLog("System", "开始运行APISequence：" + NameCN + "(" + NameEN + ")" + Thread.CurrentThread.ManagedThreadId.ToString("00"));
                                        APISequence(NameEN);
                                        System.Threading.Thread.Sleep(1000);
                                    }
                                }
                                else
                                {
                                    WiterLog("System", "TimeProjectJob.xml中[" + NameCN + "(" + NameEN + ")]配置Type錯誤,无法执行");

                                    //执行出现问题,发邮件给管理员
                                    //[ACTION] -XXXX
                                    SendActionEmail(null, "System", "System", "System配置错误", "TimeProjectJob.xml中["+ NameCN + "("+NameEN+ ")]配置Type錯誤,无法执行");
                                }
                            }
                        }

                        //PassValue.TimeProject目前所有正在执行的job,需要去掉没有在TimeProjectJob中出现的数据
                        //WiterLog("System", "正在运行数量：" + PassValue.TimeProject.Count() + "!");
                        //倒叙排列 
                        for (int i = PassValue.TimeProject.Count()-1; i >= 0; i--)
                        {
                            //是否存在于该执行的Job中
                            if (!TimeProjectJob.Contains(PassValue.TimeProject[i]))
                            {
                                //不存在 则移除
                                string TimeProjectFun = PassValue.TimeProject[i];
                                PassValue.TimeProject.Remove(TimeProjectFun);

                                WiterLog("System", "停止运行" + TimeProjectFun + "!");
                            }
                        }
                    }
                    else
                    {
                        //没有数据
                        //停止所有程序的运行 赋空值
                        PassValue.TimeProject = new List<string>();
                    }
                }

                //本轮执行完毕，赋值0，便于开启下轮启动
                OnStartStatus = 0;
                #region 多线程
                //TaskFactory taskFactory = new TaskFactory();
                //List<Task> tasks = new List<Task>();
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());
                //tasks.Add(taskFactory.StartNew());

                //tasks[0].

                //taskFactory.ContinueWhenAll(tasks.ToArray(),)
                #endregion

            }
            catch (Exception ex)
            {
                //本轮执行完毕，赋值0，便于开启下轮启动
                OnStartStatus = 0;

                EventLog.WriteEntry("我正在执行System_Time,出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                WiterLog("System", "我正在执行System_Time,出现问题：" + ex.Message + "!");
                WiterLog("System", "我正在执行System_Time,出现问题：" + ex.ToString() + "!");

                //执行出现问题,发邮件给管理员
                //[ACTION] -XXXX
                SendActionEmail(null, "System", "System", ex.Message, ex.ToString());
            }
        }

        //服务停止执行代码
        protected override void OnStop()
        {
            try
            {
                //清空执行锁定程序 便于下次执行
                PassValue.TimeProjectJob = new List<string>();

                //停止Timer执行
                TimerProjectJob.Enabled = false;
                EventLog.WriteEntry("我的服务停止");
                WiterLog("System", "我的服务停止");

            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("我正在执行停止运行,出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                WiterLog("System", "我正在执行停止运行,出现问题：" + ex.Message + "!");
                WiterLog("System", "我正在执行停止运行,出现问题：" + ex.ToString() + "!");

                //执行出现问题,发邮件给管理员
                //[ACTION] -XXXX
                SendActionEmail(null, "System", "SystemOnStop", ex.Message, ex.ToString());
            }

        }

        #region  调用API
        //开始执行异步方法
        public async void ApiInfo(string Fun)
        {
            await Task.Run(() =>
            {
                try
                {
                    //定义全局时间,用来执行时间间隔
                    DateTime Time = new DateTime();

                    //创建一个定时器
                    System.Timers.Timer TimerProject;  //计时器

                    //初始化
                    TimerProject = new System.Timers.Timer();
                    //设定定时时间  最小单位是1秒
                    TimerProject.Interval = 1000;  //设置计时器事件间隔执行时间

                    //设置定时执行方法  到达时间的时候执行事件； 
                    //TimerProject.Elapsed += new System.Timers.ElapsedEventHandler(TimerProject_Elapsed);
                    TimerProject.Elapsed += new System.Timers.ElapsedEventHandler((o, e) => TimerProject_Elapsed(o, e, Fun, TimerProject));

                    TimerProject.AutoReset = true;//设置是执行一次（false）还是一直执行(true)； 
                                                  //打开Timer执行
                    TimerProject.Enabled = true; //是否执行System.Timers.Timer.Elapsed事件；
                }
                catch (Exception ex)
                {
                    EventLog.WriteEntry("我正在执行" + Fun + "-ApiInfo,出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                    WiterLog(Fun, "我正在执行" + Fun + "-ApiInfo,出现问题：" + ex.Message + "!");
                    WiterLog(Fun, "我正在执行" + Fun + "-ApiInfo,,出现问题：" + ex.ToString() + "!");

                    //执行出现问题,发邮件给管理员
                    //[ACTION] -XXXX
                    SendActionEmail(null, Fun, Fun + "-ApiInfo", ex.Message, ex.ToString());
                }
            });
        }

        //定时程序执行方法
        private void TimerProject_Elapsed(object sender, System.Timers.ElapsedEventArgs e, string Fun, System.Timers.Timer TimerProject)
        {
            //記錄方法配置
            DataTable dt = new DataTable();
            try
            {
                //判断程序是否还需要执行
                //不存在与程序执行List中,停止运行程序
                if (!PassValue.TimeProject.Contains(Fun))
                {
                    WiterLog(Fun, "停止运行ApiInfo" + Fun + "-Time!");

                    //将标注每个Fun准确的执行时间中的数据清除，便于下次执行
                    PassValue.ExecutionStatusByTime.Remove(Fun);

                    //将今天已经执行过的周期记录也清除掉，便于下次执行
                    PassValue.ExecutionStatus.Remove(Fun);

                    //清除执行锁定程序 便于下次执行
                    PassValue.TimeProjectJob.Remove(Fun);

                    TimerProject.Enabled = false;
                    return;
                }

                //WiterLog(Fun, "我正在执行ApiInfo" + Fun + "-Time!");

                //获取配置文件根目录
                string basePath = File.LocalConfig();

                //拼接System应该存在的配置文件路径
                string ConfigPath = Path.Combine(basePath, Fun);

                //检查目录是否存在,不存在则创建
                if (File.DirectoryIsExist(ConfigPath))
                {
                    //获取Xml配置的信息
                     dt = File.GetXmlInfo(ConfigPath, Fun + "_Setup.xml");
                    //判断是否有数据
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        DateTime NowTime = DateTime.Now;
                        //无论有多少行，只取第一行
                        //开始时间 <!--开始时间 yyyy-MM-dd HH:mm  24小时制 没有则留空 与结束时间无需成对-->
                        string StartTime = dt.Rows[0]["StartTime"].ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(StartTime))
                        {
                            //判断是否到达约定时间  当前时间小于等于开始时间的时候 直接结束
                            if (NowTime < Convert.ToDateTime(StartTime))
                            {
                                WiterLog(Fun, Fun + "-Time,当前时间："+ NowTime .ToString()+ "还没有到达开始时间："+StartTime+",不能开始执行程序");
                                return;
                            }
                        }

                        //结束时间 <!--结束时间 yyyy-MM-dd  HH:mm 24小时制 没有则留空 与开始时间无需成对--->
                        string EndTime = dt.Rows[0]["EndTime"].ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(EndTime))
                        {
                            //判断是否到达约定时间  当前时间大于结束时间的时候 直接结束
                            if (NowTime > Convert.ToDateTime(EndTime))
                            {
                                WiterLog(Fun, Fun + "-Time,当前时间：" + NowTime.ToString() + "已经超过结束时间：" + EndTime + ",不能开始执行程序");
                                return;
                            }
                        }

                        //判断执行类型 <!--执行类型 0:按时间间隔 1:按周期-->
                        string ExecutionStatus = dt.Rows[0]["ExecutionStatus"].ToString().Trim();
                        if (ExecutionStatus == "0")//0:按时间间隔
                        {
                            //获取时间间隔
                            string IntervalsTime = dt.Rows[0]["IntervalsTime"].ToString().Trim();
                            if (IntervalsTime == "")
                            {
                                WiterLog(Fun, Fun + "_Setup.xml中IntervalsTime配置错误,无法执行。请按照说明调整(时间间隔 单位:秒)。");

                                SendActionEmail(dt, Fun, Fun + "-Time", "配置错误", Fun + "_Setup.xml中IntervalsTime配置错误,无法执行。请按照说明调整(时间间隔 单位:秒)。");
                                return;
                            }
                            else
                            {
                                //判断是第一次进方法就给全局时间赋当前时间
                                if (!PassValue.ExecutionStatusByTime.ContainsKey(Fun))
                                {
                                    //添加准确执行时间到数组上
                                    PassValue.ExecutionStatusByTime.Add(Fun, DateTime.Now);
                                    //执行 全局时间赋值在主程序里面
                                    ApiInfoLoadData(Fun, dt, "时间间隔,第一次执行");  //第一次执行
                                    return;
                                }
                                else
                                {
                                    DateTime Time = PassValue.ExecutionStatusByTime[Fun];
                                    //设置时间相等时 执行不进去,设置时间小于 执行一次 并马上赋值Time为目前时间
                                    if (Time.AddSeconds(double.Parse(IntervalsTime)) <= DateTime.Now)
                                    {
                                        //执行
                                        ApiInfoLoadData(Fun, dt,"时间间隔,第二次执行"); //第二次执行
                                        return;
                                    }
                                }
                            }
                        }
                        else if (ExecutionStatus == "1")//1:按周期
                        {
                            //周期选择 ： EveryDay EveryWeek
                            string CycltType = dt.Rows[0]["CycltType"].ToString().Trim();

                            //分析周期
                            if (CycltType.ToLower() == "everyweek")
                            {
                                //星期几 ： Monday  Tuesday Wednesday Thursday  Friday  Saturday  Sunday
                                string DayOfWeek = dt.Rows[0]["DayOfWeek"].ToString().Trim();
                                //判断星期几是否符合
                                if (DayOfWeek.ToLower() == DateTime.Today.DayOfWeek.ToString().ToLower())
                                {
                                    //获取 特定时间：HH:mm 24小时制 不包含秒
                                    string SpecificTime = dt.Rows[0]["SpecificTime"].ToString().Trim();
                                    //去掉秒,执行准确
                                    if (Convert.ToInt32(Convert.ToDateTime(SpecificTime).ToString("HHmm")) == Convert.ToInt32(DateTime.Now.ToString("HHmm")))
                                    {
                                        //今天是否有执行过,未执行过才能执行，否则这一分钟会执行60次
                                        if (!PassValue.ExecutionStatus.Contains(Fun))
                                        {
                                            //将当前执行方法添加到List中,每天晚上00:00 清空这个List
                                            PassValue.ExecutionStatus.Add(Fun);
                                            //执行方法
                                            ApiInfoLoadData(Fun, dt, "周期");//周期
                                            return;
                                        }
                                    }
                                }
                            }
                            else if (CycltType.ToLower() == "everyday")
                            {
                                //获取 特定时间：HH:mm 24小时制 不包含秒
                                string SpecificTime = dt.Rows[0]["SpecificTime"].ToString().Trim();
                                //去掉秒,执行准确
                                if (Convert.ToInt32(Convert.ToDateTime(SpecificTime).ToString("HHmm")) == Convert.ToInt32(DateTime.Now.ToString("HHmm")))
                                {
                                    //今天是否有执行过,未执行过才能执行，否则这一分钟会执行60次
                                    if (!PassValue.ExecutionStatus.Contains(Fun))
                                    {
                                        //将当前执行方法添加到List中,每天晚上00:00 清空这个List
                                        PassValue.ExecutionStatus.Add(Fun);
                                        //执行方法
                                        ApiInfoLoadData(Fun, dt,"周期");//周期
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                WiterLog(Fun, "我正在执行ApiInfo_" + Fun + "_Time,配置文件：CycltType 配置错误,请按照说明调整（周期选择 ： EveryDay EveryWeek）!");

                                SendActionEmail(dt, Fun, Fun + "-Time", "配置错误", Fun + "_Setup.xml中CycltType配置错误,无法执行。请按照说明调整(周期选择:EveryDay EveryWeek)。");
                            }
                        }
                        else
                        {
                            WiterLog(Fun, "我正在执行ApiInfo_" + Fun + "_Time,配置文件：ExecutionStatus 配置错误,请按照说明调整（执行类型 0:按时间间隔 1:按周期）!");
                            SendActionEmail(dt, Fun, Fun + "-Time", "配置错误", Fun + "_Setup.xml中ExecutionStatus配置错误,无法执行。请按照说明调整(执行类型 0:按时间间隔 1:按周期)。");
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("我正在执行ApiInfo_" + Fun + "_Time,出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                WiterLog(Fun, "我正在执行ApiInfo_" + Fun + "_Time,出现问题：" + ex.Message + "!");
                WiterLog(Fun, "我正在执行ApiInfo_" + Fun + "_Time,出现问题：" + ex.ToString() + "!");

                //执行出现问题,发邮件给管理员
                //[ACTION] -XXXX
                SendActionEmail(dt, Fun, Fun + "-Time", ex.Message, ex.ToString());
            }
        }


        /// <summary>
        /// 调用Api时候的主程序
        /// </summary>
        /// <param name="Fun">方法名称</param>
        /// <param name="dt">传过来配置信息,便于继续往下执行</param>
        /// <param name="num">是第一次执行,还是第二次执行</param>
        public void ApiInfoLoadData(string Fun, DataTable dt,string Info)
        {
            //判断当前方法是否在运行 还在运行则不允许继续执行
            if (PassValue.TimeProjectJob.Contains(Fun))
            {
                WiterLog(Fun, "程序还未处理完成，锁定执行");
                return;
            }

            string ApiUrl = dt.Rows[0]["ApiUrl"].ToString().Trim();
            if (string.IsNullOrWhiteSpace(ApiUrl))
            {
                WiterLog(Fun, Fun + "_Setup.xml中ApiUrl配置错误,无法执行。请按照说明调整(Api 请求地址)。");

                SendActionEmail(dt, Fun, Fun + "-ApiInfo", "配置错误", Fun + "_Setup.xml中ApiUrl配置错误,无法执行。请按照说明调整(Api 请求地址)。");
                return;
            }

            //获取是否需要安全验证，是则必须填写验证密钥
            string Verification= dt.Rows[0]["Verification"].ToString().Trim();
            string AuthenticationKey = dt.Rows[0]["AuthenticationKey"].ToString().Trim();
            if (Verification.ToLower()=="true")
            {
                if (AuthenticationKey == "")
                {
                    WiterLog(Fun, Fun + "_Setup.xml中AuthenticationKey配置错误,无法执行。请按照说明调整(请求安全验证之密钥)。");

                    SendActionEmail(dt, Fun, Fun + "-ApiInfo", "配置错误", Fun + "_Setup.xml中AuthenticationKey配置错误,无法执行。请按照说明调整(请求安全验证之密钥)。");
                    return;
                }
                else
                {
                    //拼接新的Url
                    string time = LocalFile.GetTimeStampByMilliseconds();
                    ApiUrl += "?Timestamp=" + time + "&SyncKey=" + LocalFile.GetSha1(AuthenticationKey + time);
                }
            }

            //当前时间赋值给全局时间 方便下次循环
            //判断是否存在数据在全局变量里面
            if (PassValue.ExecutionStatusByTime.ContainsKey(Fun))
            {
                //PassValue.ExecutionStatusByTime.Remove(Fun);
                //存在,修改
                PassValue.ExecutionStatusByTime[Fun]= DateTime.Now;
            }
            else
            {
                //不存在添加
                PassValue.ExecutionStatusByTime.Add(Fun, DateTime.Now);
            }

            //将当前方法写入List  避免重复执行
            PassValue.TimeProjectJob.Add(Fun);

            WiterLog(Fun, Info + ",程序开始执行!");
            try
            {
                string AliResult = HttpRequest.HttpGet(ApiUrl, "", 2 * 60 * 60 * 1000); //设置4个小时的超时时间

                WiterLog(Fun, "返回数据：" + AliResult);

                #region 是否需要在执行成功的清空下发送邮件

                SendInfoEmail(dt,Fun, Fun + "-LoadData", "執行成功", AliResult);

                #endregion

                //执行完成后，从List中删除,便于执行第二次方法。
                PassValue.TimeProjectJob.Remove(Fun);
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("我正在执行LoadData_" + Fun + "_Time,出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                WiterLog(Fun, "我正在执行LoadData_" + Fun + "_Time,出现问题：" + ex.Message + "!");
                WiterLog(Fun, "我正在执行LoadData_" + Fun + "_Time,出现问题：" + ex.ToString() + "!");

                //执行完成后，从List中删除,便于执行第二次方法。
                PassValue.TimeProjectJob.Remove(Fun);

                //發送郵件
                SendActionEmail(null, Fun, Fun + "-LoadData", ex.Message, ex.ToString());
            }
        }

        #endregion

        #region  顺序执行（API+存储过程）
        //开始执行异步方法
        public async void APISequence(string Fun)
        {
            await Task.Run(() =>
            {
                try
                {

                    //创建一个定时器
                    System.Timers.Timer TimerProject;  //计时器

                    //初始化
                    TimerProject = new System.Timers.Timer();
                    //设定定时时间  最小单位是1秒
                    TimerProject.Interval = 1000;  //设置计时器事件间隔执行时间

                    //设置定时执行方法  到达时间的时候执行事件； 
                    //TimerProject.Elapsed += new System.Timers.ElapsedEventHandler(TimerProject_Elapsed);
                    TimerProject.Elapsed += new System.Timers.ElapsedEventHandler((o, e) => TimerProjectAPISequence_Elapsed(o, e, Fun, TimerProject));

                    TimerProject.AutoReset = true;//设置是执行一次（false）还是一直执行(true)； 
                                                  //打开Timer执行
                    TimerProject.Enabled = true; //是否执行System.Timers.Timer.Elapsed事件；
                }
                catch (Exception ex)
                {
                    EventLog.WriteEntry("我正在执行" + Fun + ",出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                    WiterLog(Fun, "我正在执行" + Fun + ",出现问题：" + ex.Message + "!");
                    WiterLog(Fun, "我正在执行" + Fun + ",出现问题：" + ex.ToString());

                    //执行出现问题,发邮件给管理员
                    //[ACTION] -XXXX
                    SendActionEmail(null, Fun, Fun + "-APISequence", ex.Message, ex.ToString());
                }
            });
        }

        //定时程序执行方法
        private void TimerProjectAPISequence_Elapsed(object sender, System.Timers.ElapsedEventArgs e, string Fun, System.Timers.Timer TimerProject)
        {
            //記錄方法配置
            DataTable dt = new DataTable();
            try
            {
                //判断程序是否还需要执行
                //不存在与程序执行List中,停止运行程序
                if (!PassValue.TimeProject.Contains(Fun))
                {
                    WiterLog(Fun, "停止运行APISequence" + Fun + "-Time!");

                    //将标注每个Fun准确的执行时间中的数据清除，便于下次执行
                    PassValue.ExecutionStatusByTime.Remove(Fun);

                    //将今天已经执行过的周期记录也清除掉，便于下次执行
                    PassValue.ExecutionStatus.Remove(Fun);

                    //清除执行锁定程序 便于下次执行
                    PassValue.TimeProjectJob.Remove(Fun);

                    TimerProject.Enabled = false;
                    return;
                }

                //WiterLog(Fun, "我正在执行APISequence" + Fun + "-Time!");


                //获取配置文件根目录
                string basePath = File.LocalConfig();

                //拼接System应该存在的配置文件路径
                string ConfigPath = Path.Combine(basePath, Fun);

                //检查目录是否存在,不存在则创建
                if (File.DirectoryIsExist(ConfigPath))
                {
                    //获取Xml配置的信息
                    dt = File.GetXmlInfo(ConfigPath, Fun + "_Setup.xml");
                    //判断是否有数据
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        DateTime NowTime = DateTime.Now;
                        //无论有多少行，只取第一行
                        //开始时间 <!--开始时间 yyyy-MM-dd HH:mm  24小时制 没有则留空 与结束时间无需成对-->
                        string StartTime = dt.Rows[0]["StartTime"].ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(StartTime))
                        {
                            //判断是否到达约定时间  当前时间小于等于开始时间的时候 直接结束
                            if (NowTime < Convert.ToDateTime(StartTime))
                            {
                                WiterLog(Fun, Fun + "-Time,当前时间：" + NowTime.ToString() + "还没有到达开始时间：" + StartTime + ",不能开始执行程序");
                                return;
                            }
                        }

                        //结束时间 <!--结束时间 yyyy-MM-dd  HH:mm 24小时制 没有则留空 与开始时间无需成对--->
                        string EndTime = dt.Rows[0]["EndTime"].ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(EndTime))
                        {
                            //判断是否到达约定时间  当前时间大于结束时间的时候 直接结束
                            if (NowTime > Convert.ToDateTime(EndTime))
                            {
                                WiterLog(Fun, Fun + "-Time,当前时间：" + NowTime.ToString() + "已经超过结束时间：" + EndTime + ",不能开始执行程序");
                                return;
                            }
                        }

                        //判断执行类型 <!--执行类型 0:按时间间隔 1:按周期-->
                        string ExecutionStatus = dt.Rows[0]["ExecutionStatus"].ToString().Trim();
                        if (ExecutionStatus == "0")//0:按时间间隔
                        {
                            //获取时间间隔
                            string IntervalsTime = dt.Rows[0]["IntervalsTime"].ToString().Trim();
                            if (IntervalsTime == "")
                            {
                                WiterLog(Fun, Fun + "_Setup.xml中IntervalsTime配置错误,无法执行。请按照说明调整(时间间隔 单位:秒)。");

                                SendActionEmail(dt, Fun, Fun + "-Time", "配置错误", Fun + "_Setup.xml中IntervalsTime配置错误,无法执行。请按照说明调整(时间间隔 单位:秒)。");

                                return;
                            }
                            else
                            {
                                //判断是第一次进方法就给全局时间赋当前时间
                                if (!PassValue.ExecutionStatusByTime.ContainsKey(Fun))
                                {
                                    //添加准确执行时间到数组上
                                    PassValue.ExecutionStatusByTime.Add(Fun, DateTime.Now);

                                    //执行 全局时间赋值在主程序里面
                                    APISequenceLoadData(Fun, dt);
                                    return;
                                }
                                else
                                {
                                    DateTime Time = PassValue.ExecutionStatusByTime[Fun];
                                    //设置时间相等时 执行不进去,设置时间小于 执行一次 并马上赋值Time为目前时间
                                    if (Time.AddSeconds(double.Parse(IntervalsTime)) <= DateTime.Now)
                                    {
                                        //执行
                                        APISequenceLoadData(Fun, dt);
                                        return;
                                    }
                                }
                            }
                        }
                        else if (ExecutionStatus == "1")//1:按周期
                        {
                            //周期选择 ： EveryDay EveryWeek
                            string CycltType = dt.Rows[0]["CycltType"].ToString().Trim();

                            //分析周期
                            if (CycltType.ToLower() == "everyweek")
                            {
                                //星期几 ： Monday  Tuesday Wednesday Thursday  Friday  Saturday  Sunday
                                string DayOfWeek = dt.Rows[0]["DayOfWeek"].ToString().Trim();
                                //判断星期几是否符合
                                if (DayOfWeek.ToLower() == DateTime.Today.DayOfWeek.ToString().ToLower())
                                {
                                    //获取 特定时间：HH:mm 24小时制 不包含秒
                                    string SpecificTime = dt.Rows[0]["SpecificTime"].ToString().Trim();
                                    //去掉秒,执行准确
                                    if (Convert.ToInt32(Convert.ToDateTime(SpecificTime).ToString("HHmm")) == Convert.ToInt32(DateTime.Now.ToString("HHmm")))
                                    {
                                        //今天是否有执行过,未执行过才能执行，否则这一分钟会执行60次
                                        if (!PassValue.ExecutionStatus.Contains(Fun))
                                        {
                                            //将当前执行方法添加到List中,每天晚上00:00 清空这个List
                                            PassValue.ExecutionStatus.Add(Fun);
                                            //执行方法
                                            APISequenceLoadData(Fun, dt);
                                            return;
                                        }
                                    }
                                }
                            }
                            else if (CycltType.ToLower() == "everyday")
                            {
                                //获取 特定时间：HH:mm 24小时制 不包含秒
                                string SpecificTime = dt.Rows[0]["SpecificTime"].ToString().Trim();
                                //去掉秒,执行准确
                                if (Convert.ToInt32(Convert.ToDateTime(SpecificTime).ToString("HHmm")) == Convert.ToInt32(DateTime.Now.ToString("HHmm")))
                                {
                                    //今天是否有执行过,未执行过才能执行，否则这一分钟会执行60次
                                    if (!PassValue.ExecutionStatus.Contains(Fun))
                                    {
                                        //将当前执行方法添加到List中,每天晚上00:00 清空这个List
                                        PassValue.ExecutionStatus.Add(Fun);
                                        //执行方法
                                        APISequenceLoadData(Fun, dt);
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                WiterLog(Fun, "我正在执行APISequence_" + Fun + "_Time,配置文件：CycltType 配置错误,请按照说明调整（周期选择 ： EveryDay EveryWeek）!");

                                SendActionEmail(dt, Fun, Fun + "-Time", "配置错误", Fun + "_Setup.xml中CycltType配置错误,无法执行。请按照说明调整(周期选择:EveryDay EveryWeek)。");
                            }
                        }
                        else
                        {
                            WiterLog(Fun, "我正在执行APISequence_" + Fun + "_Time,配置文件：ExecutionStatus 配置错误,请按照说明调整（执行类型 0:按时间间隔 1:按周期）!");

                            SendActionEmail(dt, Fun, Fun + "-Time", "配置错误", Fun + "_Setup.xml中ExecutionStatus配置错误,无法执行。请按照说明调整(执行类型 0:按时间间隔 1:按周期)。");
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("我正在执行APISequence_" + Fun + "_Time,出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                WiterLog(Fun, "我正在执行APISequence_" + Fun + "_Time,出现问题：" + ex.Message + "!");
                WiterLog(Fun, "我正在执行APISequence_" + Fun + "_Time,出现问题：" + ex.ToString());

                //执行出现问题,发邮件给管理员
                //[ACTION] -XXXX
                SendActionEmail(dt, Fun, Fun + "-Time", ex.Message, ex.ToString());
            }
        }

        /// <summary>
        /// 顺序执行时候的主程序
        /// </summary>
        /// <param name="Time">接受传入的时间,返回赋值之后的时间 ref不赋值也行</param>
        /// <param name="Fun">方法名称</param>
        public void APISequenceLoadData(string Fun, DataTable dt)
        {
            //判断当前方法是否在运行 还在运行则不允许继续执行
            if (PassValue.TimeProjectJob.Contains(Fun))
            {
                WiterLog(Fun, "程序还未处理完成，锁定执行");
                return;
            }

            //当前时间赋值给全局时间 方便下次循环
            //判断是否存在数据在全局变量里面
            if (PassValue.ExecutionStatusByTime.ContainsKey(Fun))
            {
                //PassValue.ExecutionStatusByTime.Remove(Fun);
                //存在,修改
                PassValue.ExecutionStatusByTime[Fun] = DateTime.Now;
            }
            else
            {
                //不存在添加
                PassValue.ExecutionStatusByTime.Add(Fun, DateTime.Now);
            }

            //将当前方法写入List  避免重复执行
            PassValue.TimeProjectJob.Add(Fun);

            WiterLog(Fun, "程序开始执行!");
            try
            {

                string AliResult = "";

                //获取配置文件根目录
                string basePath = File.LocalConfig();

                //拼接System应该存在的配置文件路径  打开指定的项目的文件夹
                string ConfigPath = Path.Combine(basePath , Fun);

                //检查目录是否存在,不存在则创建
                if (File.DirectoryIsExist(ConfigPath))
                {
                    //获取Xml配置的信息
                    DataTable SequenceDT = File.GetXmlInfo(ConfigPath, Fun + "_Sequence.xml");
                    //判断是否有数据
                    if (SequenceDT != null && SequenceDT.Rows.Count > 0)
                    {
                        for (int i = 0; i < SequenceDT.Rows.Count; i++)
                        {
                            string res = "";
                            WiterLog(Fun, "开始循环");
                            WiterLog(Fun, "执行第" + (i + 1) + "个" + SequenceDT.Rows[i][0].ToString() + " ：" + SequenceDT.Rows[i][1].ToString());
                            if (SequenceDT.Rows[i][0].ToString() == "存储过程")
                            {
                                string ConnStr = dt.Rows[0]["ConnStr"].ToString().Trim();
                                if (string.IsNullOrWhiteSpace(ConnStr))
                                {
                                   
                                    WiterLog(Fun, "我正在执行APISequence_" + Fun + "_Time," + Fun + "_Sequence.xml中ConnStr配置错误,无法执行。请按照说明调整(数据库链接字符串  Server = xxxx.xxx.xx; Database = Test; User ID = sa; Password = 123;)。");

                                    SendActionEmail(dt, Fun, Fun + "-APISequence", "配置错误", Fun + "_Sequence.xml中ConnStr配置错误,无法执行。请按照说明调整(数据库链接字符串  Server=xxxx.xxx.xx;Database=Test;User ID=sa;Password=123;)。");

                                    //执行完成后，从List中删除,便于执行第二次方法。
                                    PassValue.TimeProjectJob.Remove(Fun);

                                    return;
                                }

                                //写入参数
                                DataTable dtSql = SQLHelper.ExecuteDataTable(ConnStr, CommandType.StoredProcedure, SequenceDT.Rows[i][1].ToString(), null);
                                if (dtSql != null && dtSql.Rows.Count > 0)
                                {
                                    res += dtSql.Rows[0][0].ToString();
                                }
                            }
                            else if (SequenceDT.Rows[i][0].ToString() == "API地址")
                            {
                  

                                //获取是否需要安全验证，是则必须填写验证密钥
                                string Verification = dt.Rows[0]["Verification"].ToString().Trim();
                                string AuthenticationKey = dt.Rows[0]["AuthenticationKey"].ToString().Trim();
                                string ApiUrl = SequenceDT.Rows[i][1].ToString().Trim();

                                if (string.IsNullOrWhiteSpace(ApiUrl))
                                {
                                    WiterLog(Fun, Fun + "_Setup.xml中ApiUrl配置错误,无法执行。请按照说明调整(Api 请求地址)。");

                                    SendActionEmail(dt, Fun, Fun + "-APISequence", "配置错误", Fun + "_Sequence.xml中ApiUrl配置错误,无法执行。请按照说明调整(Api 请求地址)。");
                                    return;
                                }

                                if (Verification.ToLower() == "true")
                                {
                                    if (AuthenticationKey == "")
                                    {
                                        WiterLog(Fun, "我正在执行APISequence_" + Fun + "_Time,配置文件：AuthenticationKey 配置错误,请按照说明调整（请求安全验证之密钥）!");

                                        SendActionEmail(dt, Fun, Fun + "-APISequence", "配置错误", Fun + "_Sequence.xml中AuthenticationKey配置错误,无法执行。请按照说明调整(请求安全验证之密钥)。");

                                        //执行完成后，从List中删除,便于执行第二次方法。
                                        PassValue.TimeProjectJob.Remove(Fun);

                                        return;
                                    }
                                    else
                                    {
                                        //拼接新的Url
                                        string time = LocalFile.GetTimeStampByMilliseconds();
                                        ApiUrl += "?Timestamp=" + time + "&SyncKey=" + LocalFile.GetSha1(AuthenticationKey + time);
                                    }
                                }

                                res = HttpRequest.HttpGet(ApiUrl, "", 2 * 60 * 60 * 1000);//设置2个小时的超时时间
                            }
                            AliResult += "序号：" + (i + 1) + SequenceDT.Rows[i][0].ToString() + ":" + SequenceDT.Rows[i][1].ToString() + "；返回数据：" + res;

                            WiterLog(Fun, "序号：" + (i + 1) + SequenceDT.Rows[i][0].ToString() + ":" + SequenceDT.Rows[i][1].ToString() + "；返回数据：" + res);
                        }
                    }
                }

                //WiterLog(Fun, "返回数据：" + AliResult);


                #region 是否需要在执行成功的情况下发送邮件
                SendInfoEmail(dt, Fun, Fun + "-LoadData", "執行成功", AliResult);
                #endregion

                //执行完成后，从List中删除,便于执行第二次方法。
                PassValue.TimeProjectJob.Remove(Fun);
            }
            catch (Exception ex)
            {
                WiterLog(Fun, "发生错误：" + ex.Message);
                WiterLog(Fun, "发生错误：" + ex.ToString());

                //發送郵件
                SendActionEmail(null, Fun, Fun + "-LoadData", ex.Message, ex.ToString());

                //执行完成后，从List中删除,便于执行第二次方法。
                PassValue.TimeProjectJob.Remove(Fun);
            }
        }
        #endregion

        #region  执行存储过程
        //Stored procedure
        #endregion


        private static object o = new object();

        #region 写入日志  最好异步写,避免线程阻塞
        /// <summary>
        /// 写入日志  最好异步写,避免线程阻塞
        /// </summary>
        /// <param name="Fun">方法名</param>
        /// <param name="Info">写入消息</param>
        public void WiterLog(string Fun, string Info)
        {
            //await Task.Run(() =>
            //{
            try
            {
                //写入日志信息表 用文件夹分组
                string fileName = DateTime.Now.ToString("yyyyMMdd") + @".log";

                //创建目录
                string basePath = File.LocalLog();

                //按月份分组日志
                string LogPath = Path.Combine(Path.Combine(basePath, Fun), DateTime.Now.ToString("yyyyMM"));

                //检查目录是否存在,不存在则创建
                if (File.DirectoryIsExist(LogPath))
                {
                    //文件与文件名组合在一起
                    string FilePath = Path.Combine(LogPath, fileName);

                    //EventLog.WriteEntry("地址：" + FilePath);//在系统事件查看器里的应用程序事件里来源的描述


                    //先写好需要写到日志文件中的数据
                    string log = "[" + DateTime.Now.ToString() + "]       " + Info + "\r\n";
                    ThreadPool.QueueUserWorkItem(new WaitCallback(obj =>//线程池，在有线程池线程变得可用时执行
                    {
                        lock (o)
                        {
                            //判断文件是否已经存在.
                            if (File.FileIsExist(FilePath))
                            {
                                using (var sw = new StreamWriter(FilePath, true))
                                {
                                    var sb = new StringBuilder();
                                    sb.Append(log);
                                    //sb.AppendLine(Environment.NewLine);
                                    sw.Write(sb.ToString());
                                }
                            }
                        }
                    }));

                    ////判断文件是否已经存在.
                    //if (File.FileIsExist(FilePath))
                    //{
                    //   System.IO.File.AppendAllText(FilePath, "[" + DateTime.Now.ToString() + "]       " + Info + "\r\n");

                    //    //FileInfo FI = new FileInfo(FilePath);
                    //    //StreamWriter SW = FI.AppendText();
                    //    //SW.WriteLine("[" + DateTime.Now.ToString() + "]       " + Info);
                    //    ////SW.Write("[" + DateTime.Now.ToString() + "]       " + Info + "\r\n");
                    //    //SW.Flush();
                    //    //SW.Close();

                    //    ////关闭后，需要手动清理一下资源
                    //    //SW.Dispose();
                    //    //SW = null;
                    //}
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("我正在执行WiterLog_" + Fun + ",出现问题：" + ex.Message + "!");//在系统事件查看器里的应用程序事件里来源的描述
                WiterLog(Fun, "我正在执行WiterLog_" + Fun + ",出现问题：" + ex.Message + "!");

                WiterLog("WiterLog", "我正在执行WiterLog_" + Fun + ",出现问题：" + ex.Message + "!");

                //發送郵件
                SendActionEmail(null, Fun, Fun + "-WiterLog", ex.Message, ex.ToString());
                PassValue.TimeProjectJob.Remove(Fun);
            }
        }
        #endregion

        #region 发送邮件
        /// <summary>
        /// 程序运行错误,发送紧急邮件
        /// [ACTION] - XXXX
        /// </summary>
        public void SendActionEmail(DataTable dt, string Fun,string FunDetails, string Message,string Details)
        {
            string MailTo = "";
            if (dt != null && dt.Rows.Count > 0)
            {
                MailTo = dt.Rows[0]["MailTo"].ToString().Trim();
            }

            MailTo = MailTo.Trim();
            //判斷fun是否配置了發送地址
            if (MailTo != "")
            {
                if (MailTo.Substring(MailTo.Length - 1, 1).LastIndexOf(";") > 0)
                {
                    MailTo = MailTo.Substring(0, MailTo.Length - 1);
                }

                MailTo += ";" + SendTo;
            }
            else
            {
                MailTo = SendTo;
            }
            
            //收件人为空 不发送邮件
            if (MailTo.Trim() != "")
            {
                WiterLog(Fun, "收件人为：" + MailTo.Trim());
                WiterLog("Mail", FunDetails+"收件人为：" + MailTo.Trim());
                //附件
                string[] file = { };

                //主题
                string Subject = " [ACTION]-["+ProgramName+"]-["+ FunDetails + "]-["+ Message + "]";

                //正文
                //string Body = "<table> <tr style=\"height: 40px;\"><td style=\"width: 80px;text-align: left;vertical-align:top;\">Dear All:</td><td style=\"width: 80px;text-align: left;vertical-align:top;\"></td style=\"text-align: left;vertical-align:top;\"><td></td></tr> <tr style=\"height: 40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\">" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 程序出現異常情況，請緊急處理！！！</td></tr> <tr style=\"height: 40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\">異常情況描述:</td><td style=\"text-align: left;vertical-align:top;\">" + Message + "</td></tr> <tr style=\"height: 40px;\"><td></td><td style=\"text-align: left;vertical-align:top;\">異常情況詳情:</td><td style=\"text-align: left;vertical-align:top;\">" + Details + "</td></tr></table>";

                string Body = "<table>";
                Body += " <tr style=\"height: 40px;\"><td style=\"width:80px;text-align:left;vertical-align:top;\">Dear All:</td><td style=\"width:80px;text-align:left;vertical-align:top;\"></td style=\"text-align:left;vertical-align:top;\"><td></td></tr>";
                Body += "<tr style=\"height:40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\"> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 程序出現異常情況，請緊急處理！！！</td></tr>";
                Body += "<tr style=\"height:40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\">異常情況描述:" + Message + " </td></tr>";
                Body += "<tr style=\"height:40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\">異常情況詳情:<br/>" + Details + "</td></tr></table>";

                //发送邮件
                string Res = Send.SendmailFile(MailTo.Trim(), file, Subject, Body, MailPriority.High);

                WiterLog(Fun, Res);
                WiterLog("Mail", FunDetails +","+Res);
            }
            else
            {
                WiterLog(Fun, "收件人为空,邮件无法发送!");
                WiterLog("Mail", FunDetails + "收件人为空,邮件无法发送!");
            }
        }

        /// <summary>
        /// 程序运行成功,发送邮件通知
        /// </summary>
        public void SendInfoEmail(DataTable dt, string Fun, string FunDetails, string Message, string Details)
        {
            string MailTo = "";
            if (dt != null && dt.Rows.Count > 0)
            {
                string SendMail = dt.Rows[0]["SendMail"].ToString().Trim();
                if (SendMail.ToLower() != "true")
                {
                    //不需要發送郵件
                    return;
                }
                MailTo = dt.Rows[0]["MailTo"].ToString().Trim();
            }

            MailTo = MailTo.Trim();
            
            if (InfoMessage.ToLower()== "true")
            {
                //判斷fun是否配置了發送地址
                if (MailTo != "")
                {
                    if (MailTo.Substring(MailTo.Length - 1, 1).LastIndexOf(";") > 0)
                    {
                        MailTo = MailTo.Trim();
                        MailTo = MailTo.Substring(0, MailTo.Length - 1);
                    }

                    MailTo += ";" + SendTo;
                }
                else
                {
                    MailTo = SendTo;
                }
            }

            //收件人为空 不发送邮件
            if (MailTo.Trim() != "")
            {
                WiterLog(Fun, "收件人为：" + MailTo.Trim());
                WiterLog("Mail", FunDetails + "收件人为：" + MailTo.Trim());
                //附件
                string[] file = { };

                //主题
                string Subject = " [INFO]-[" + ProgramName + "]-[" + FunDetails + "]-[" + Message + "]";

                //正文
                //string Body = "<table> <tr style=\"height: 40px;\"><td style=\"width: 80px;text-align: left;vertical-align:top;\">Dear All:</td><td style=\"width: 80px;text-align: left;vertical-align:top;\"></td style=\"text-align: left;vertical-align:top;\"><td></td></tr> <tr style=\"height: 40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\">" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 程序執行成功</td></tr> <tr style=\"height: 40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\">執行結果:</td><td style=\"text-align: left;vertical-align:top;\">" + Message + "</td></tr> <tr style=\"height: 40px;\"><td></td><td style=\"text-align: left;vertical-align:top;\">執行結果詳情:</td><td style=\"text-align: left;vertical-align:top;\">" + Details + "</td></tr></table>";

                string Body = "<table>";
                Body += " <tr style=\"height: 40px;\"><td style=\"width:80px;text-align:left;vertical-align:top;\">Dear All:</td><td style=\"width:80px;text-align:left;vertical-align:top;\"></td style=\"text-align:left;vertical-align:top;\"><td></td></tr>";
                Body += "<tr style=\"height:40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\"> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 程序執行成功</td></tr>";
                Body += "<tr style=\"height:40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\">執行結果:" + Message + " </td></tr>";
                Body += "<tr style=\"height:40px;\"><td style=\"text-align: left;vertical-align:top;\"></td><td style=\"text-align: left;vertical-align:top;\" colspan=\"2\">執行結果詳情:<br/> " + Details + "</td></tr></table>";


                //发送邮件
                string Res = Send.SendmailFile(MailTo.Trim(), file, Subject, Body,MailPriority.Normal);

                WiterLog(Fun, Res);
                WiterLog("Mail", FunDetails + "," + Res);
            }
            else
            {
                WiterLog(Fun, "收件人为空,邮件无法发送!");
                WiterLog("Mail", FunDetails + "收件人为空,邮件无法发送!");
            }
        }
        #endregion
    }
}
