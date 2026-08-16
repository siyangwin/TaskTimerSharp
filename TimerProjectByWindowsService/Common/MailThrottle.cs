using System;
using System.Collections.Generic;

namespace TimerProjectByWindowsService.Common
{
    /// <summary>
    /// ACTION告警邮件节流：同一任务同一错误在节流窗口内只发送一封，
    /// 避免配置错误等场景下每秒一封的邮件风暴。
    /// </summary>
    public static class MailThrottle
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, DateTime> LastSentAt = new Dictionary<string, DateTime>();

        /// <summary>节流窗口（分钟），同类错误只发一封</summary>
        public static int ThrottleMinutes = 60;

        /// <summary>判断key对应的邮件当前是否允许发送；允许则记录本次发送时间</summary>
        public static bool Allow(string key)
        {
            lock (Sync)
            {
                DateTime now = DateTime.Now;

                //顺带清理过期记录，防止字典无限增长
                if (LastSentAt.Count > 500)
                {
                    List<string> expired = new List<string>();
                    foreach (var kv in LastSentAt)
                    {
                        if ((now - kv.Value).TotalMinutes >= ThrottleMinutes) expired.Add(kv.Key);
                    }
                    foreach (string k in expired) LastSentAt.Remove(k);
                }

                DateTime last;
                if (LastSentAt.TryGetValue(key, out last) && (now - last).TotalMinutes < ThrottleMinutes)
                {
                    return false;
                }
                LastSentAt[key] = now;
                return true;
            }
        }
    }
}
