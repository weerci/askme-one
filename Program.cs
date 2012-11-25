using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading;

namespace Askme
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                // Можно запустить только один экзепляр приложения
                bool onlyInstance;
                Mutex mtx = new Mutex(true, "53d2c76a-d57f-4a0d-ad73-3c8e81d6a04e", out onlyInstance);
                
                if (onlyInstance)
                    Application.Run(new FormMain());
                Log.ToLog(CommonString.ApplicationStop);
            }
            catch (Exception err)
            {
                switch (err.Message)
                {
                    case ErrorCode.NOT_INI_SYSPROP:
                        MessageBox.Show(ErrorCode.NOT_INI_SYSPROP, Application.ProductName );
                        break;
                    default:
                        MessageBox.Show(ErrorCode.UNKNOWN_ERROR, Application.ProductName );
                        break;
                }
                Log.ToLog(err.Message);
                Application.Exit();
            }
        }
    }
}
