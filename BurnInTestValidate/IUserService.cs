using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BurnInTestValidate
{
    public interface IUserService
    {
        int inserthistory(PassmarkHistory objHistory);
        DataTable GetProductTypes();
        DataTable GetFGNames(int productTypeId);
        string Check_Curr_Stage(string serialno, string app_id, string stage, bool boardonline = true);
        bool ValidateUser(string username, string password);
    }
           
}
