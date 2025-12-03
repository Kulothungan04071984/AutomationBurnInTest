using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        private const string ElevatedTaskName = "MyWinFormsApp_Elevated";   // <-- YOUR TASK NAME
        private const string ElevatedFlag = "--elevated";

       

        [STAThread]
        static void Main(string[] args)
        {
           
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmBurnIntest());
           
        }

     

      

     
    }
}
