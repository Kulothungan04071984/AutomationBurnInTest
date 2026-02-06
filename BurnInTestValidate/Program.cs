using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        private const string ElevatedTaskName = "MyWinFormsApp_Elevated";   // <-- YOUR TASK NAME
        private const string ElevatedFlag = "--elevated";

        public enum DatabaseType
        {
            BurnIn,
            Reporting,
            Master
        }

        [STAThread]
        static void Main(string[] args)
        {
           
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

             var host = Host.CreateDefaultBuilder()
            .ConfigureServices(services =>
            {
                // string conStr = "Server=192.168.1.146;Database=Essencore;Trusted_Connection=True;";
                var connections = new Dictionary<DatabaseType, string>
                {
                    {
                        DatabaseType.BurnIn,
                        "server=192.168.1.181;database=SFCS;user id=sa;password=Syrma@2022"
                    },
                    {
                        DatabaseType.Reporting,
                        "server=192.168.1.146;database=Barcode;user id=sa;password=syrma@123;"
                    },
                    {
                        DatabaseType.Master,
                        "server=192.168.1.146;database=Essencore;user id=sa;password=syrma@123;"
                    }
                };
                services.AddSingleton(new DbConnectionFactory(connections));
                services.AddScoped<DataManagent, DataManagent>();
                services.AddScoped<IUserService, UserService>();
                services.AddTransient<FrmBurnIntest>();
                services.AddSingleton<FrmLogin>();
                services.AddSingleton<FrmProductSelection>();
                services.AddSingleton<UserSession>();
            })
            .Build();
            Application.Run(
              host.Services.GetRequiredService<FrmLogin>()
              );


        }

     

      

     
    }
}
