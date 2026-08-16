using System;
using System.Collections.Generic;

namespace TimerProjectByWindowsService.Common
{
    /// <summary>
    /// 全局运行状态。所有集合都在内部加锁保护，外部只能通过下面的方法访问，
    /// 避免多个定时器线程并发读写 List/Dictionary 造成数据损坏。
    /// </summary>
    public static class PassValue
    {
        private static readonly object Sync = new object();

        //存储目前所有正在运行的Job 系统级别
        private static readonly List<string> TimeProject = new List<string>();

        //存储目前正在运行中的Job  避免重复运行同一方法
        private static readonly List<string> TimeProjectJob = new List<string>();

        //将今天已经执行过的 按周期的 存储起来
        private static readonly List<string> ExecutionStatus = new List<string>();

        //标注每个Fun最近一次执行完成的时间，时间间隔模式用它计算下次执行时间
        private static readonly Dictionary<string, DateTime> ExecutionStatusByTime = new Dictionary<string, DateTime>();

        #region 正在运行的Job（系统级注册表）

        /// <summary>添加运行中的Job，已存在则返回false</summary>
        public static bool AddRunningJob(string fun)
        {
            lock (Sync)
            {
                if (TimeProject.Contains(fun)) return false;
                TimeProject.Add(fun);
                return true;
            }
        }

        public static bool ContainsRunningJob(string fun)
        {
            lock (Sync) return TimeProject.Contains(fun);
        }

        public static bool RemoveRunningJob(string fun)
        {
            lock (Sync) return TimeProject.Remove(fun);
        }

        /// <summary>返回快照副本，供遍历使用</summary>
        public static List<string> GetRunningJobs()
        {
            lock (Sync) return new List<string>(TimeProject);
        }

        public static void ClearRunningJobs()
        {
            lock (Sync) TimeProject.Clear();
        }

        #endregion

        #region 执行锁（避免同一任务重复执行）

        /// <summary>原子地尝试锁定任务；已被锁定则返回false</summary>
        public static bool LockJob(string fun)
        {
            lock (Sync)
            {
                if (TimeProjectJob.Contains(fun)) return false;
                TimeProjectJob.Add(fun);
                return true;
            }
        }

        public static void UnlockJob(string fun)
        {
            lock (Sync) TimeProjectJob.Remove(fun);
        }

        public static bool IsJobLocked(string fun)
        {
            lock (Sync) return TimeProjectJob.Contains(fun);
        }

        public static bool HasLockedJobs()
        {
            lock (Sync) return TimeProjectJob.Count > 0;
        }

        public static void UnlockAllJobs()
        {
            lock (Sync) TimeProjectJob.Clear();
        }

        #endregion

        #region 今天已执行过的周期任务

        /// <summary>原子地标记任务今天已执行；之前已标记则返回false（保证同一天只触发一次）</summary>
        public static bool MarkExecutedToday(string fun)
        {
            lock (Sync)
            {
                if (ExecutionStatus.Contains(fun)) return false;
                ExecutionStatus.Add(fun);
                return true;
            }
        }

        public static void UnmarkExecutedToday(string fun)
        {
            lock (Sync) ExecutionStatus.Remove(fun);
        }

        public static void ClearExecutedToday()
        {
            lock (Sync) ExecutionStatus.Clear();
        }

        #endregion

        #region 最近一次执行完成时间（时间间隔模式）

        public static bool TryGetLastRunTime(string fun, out DateTime time)
        {
            lock (Sync) return ExecutionStatusByTime.TryGetValue(fun, out time);
        }

        public static void SetLastRunTime(string fun, DateTime time)
        {
            lock (Sync) ExecutionStatusByTime[fun] = time;
        }

        public static void RemoveLastRunTime(string fun)
        {
            lock (Sync) ExecutionStatusByTime.Remove(fun);
        }

        #endregion
    }
}
