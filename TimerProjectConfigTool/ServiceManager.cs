using System;
using System.ComponentModel;
using System.ServiceProcess;

namespace TimerProjectConfigTool
{
    public enum ServiceState
    {
        NotInstalled,
        Running,
        Stopped,
        Other
    }

    public class ServiceStatusInfo
    {
        public ServiceState State;
        public string RawStatus = "";
    }

    /// <summary>
    /// TimeProject 服务的状态查询与重启。Stop/Start 需要管理员权限。
    /// </summary>
    public static class ServiceManager
    {
        public const string ServiceName = "TimeProject";

        public static ServiceStatusInfo GetStatus()
        {
            var info = new ServiceStatusInfo();
            try
            {
                using (var sc = new ServiceController(ServiceName))
                {
                    info.RawStatus = sc.Status.ToString();
                    if (sc.Status == ServiceControllerStatus.Running) info.State = ServiceState.Running;
                    else if (sc.Status == ServiceControllerStatus.Stopped) info.State = ServiceState.Stopped;
                    else info.State = ServiceState.Other;
                }
            }
            catch (InvalidOperationException)
            {
                info.State = ServiceState.NotInstalled;
            }
            catch (Win32Exception)
            {
                info.State = ServiceState.NotInstalled;
            }
            return info;
        }

        /// <summary>
        /// Stop → 等待停止(最多60秒) → Start → 等待运行(最多60秒)。
        /// 失败时抛出带友好提示的异常。
        /// </summary>
        public static void Restart()
        {
            if (GetStatus().State == ServiceState.NotInstalled)
            {
                throw new InvalidOperationException(Lang.Get("svc.notInstalled"));
            }

            string phase = "stop";
            try
            {
                using (var sc = new ServiceController(ServiceName))
                {
                    if (sc.Status != ServiceControllerStatus.Stopped && sc.Status != ServiceControllerStatus.StopPending)
                    {
                        sc.Stop();
                    }
                    sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(60));

                    phase = "start";
                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(60));
                }
            }
            catch (System.ServiceProcess.TimeoutException tex)
            {
                string key = phase == "stop" ? "svc.stopTimeout" : "svc.startTimeout";
                throw new InvalidOperationException(Lang.Get(key), tex);
            }
            catch (Win32Exception wex)
            {
                throw new UnauthorizedAccessException(Lang.Get("svc.adminRequired"), wex);
            }
            catch (InvalidOperationException ioex)
            {
                throw new InvalidOperationException(Lang.Get("svc.restartFailed", ioex.Message), ioex);
            }
        }
    }
}
