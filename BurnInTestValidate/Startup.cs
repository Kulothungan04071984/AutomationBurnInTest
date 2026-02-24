using Microsoft.Win32;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BurnInTestValidate
{
    public static class Startup
    {
        public static void AddToWindowsStartup()
        {
            string appName = "BurnInTestValidate";
            string appPath = Application.ExecutablePath;

            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(
     @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
            {
                if (rk.GetValue(appName) == null)
                {
                    rk.SetValue(appName, appPath);
                }
            }

        }


    }
}
