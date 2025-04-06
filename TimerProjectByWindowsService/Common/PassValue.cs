using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimerProjectByWindowsService.Common
{
    public static class PassValue
    {
        //API页面直接传值  存储目前所有正在运行的Job 系统级别
        public static List<string> TimeProject = new List<string>();


        //Api 存储目前正在运行中的Job  避免重复运行同一方法
        public static List<string> TimeProjectJob = new List<string>();

        //Api 将今天已经执行过的 按周期的 存储起来
        public static List<string> ExecutionStatus = new List<string>();

        //API 标注每个Fun准确的执行时间，便于时间间隔的时候去取下次执行时间
        public static Dictionary<string, DateTime> ExecutionStatusByTime= new Dictionary<string, DateTime>();
    }
}
