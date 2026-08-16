using System;
using System.Windows.Forms;

namespace TimerProjectConfigTool
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) =>
                MessageBox.Show(e.Exception.ToString(), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                MessageBox.Show(e.ExceptionObject.ToString(), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);

            Lang.Load(SettingsStore.LoadLanguage());
            Application.Run(new MainForm());
        }
    }
}
