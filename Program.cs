using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Linq;
using WFExceptions;

namespace Askme
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                // Можно запустить только один экзепляр приложения
                int c = Process.GetProcesses().Where( n => n.ProcessName == Application.ProductName).ToArray().Count();
                if (c == 0 || c == 1)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.ThreadException += new ThreadExceptionEventHandler(OnThreadException);
                    AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

                    Application.Run(new FormMain());
                    Log.ToLog(CommonString.ApplicationStop);
                }
                else
                {
                    Application.Exit();
                }

            }
            catch (Exception err)
            {
                try
                {
                    Log.ToLog(err.Message);
                }
                finally
                {
                    Application.Exit();
                }
            }
        }

        private static void OnThreadException(object sender, ThreadExceptionEventArgs t)
        {
            WFException.HandleError((Exception)t.Exception);
        }
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            WFException.HandleError((Exception)e.ExceptionObject);
        }

    }
}
