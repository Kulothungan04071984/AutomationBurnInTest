using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    public class UserService : IUserService
    {
        public readonly DataManagent _repo;
        public UserService(DataManagent repo)
        {
            _repo = repo;
        }

        public int inserthistory(PassmarkHistory objHistory)
        {
            return _repo.inserthistory(objHistory);
        }

        public DataTable GetProductTypes() => _repo.GetProductTypes();
        public DataTable GetFGNames(int productTypeId) => _repo.GetFGNames(productTypeId);

        public Task<services> Check_Curr_Stage(string serialno, string app_id, string stage, bool boardonline = true)
        {
            return _repo.Check_Curr_Stage(serialno, app_id, stage, boardonline);
        }

        public bool ValidateUser(string username, string password)
        {
            return _repo.ValidateUser(username, password);
        }

        public Task<FgDetails> GetFgDetails(int productId, string cusNumber)
        {
            return _repo.GetFgDetails(productId, cusNumber);
        }

        public string SQL_Upload(string PcbSno, string CusSno, bool boardfail, string Result_Remarks, string[] stages)
        {
            return _repo.SQL_Upload(PcbSno, CusSno, boardfail, Result_Remarks , stages);
        }

        public Task<string[]> startchecksfcs(string FgNumber)
        {
            return _repo.startchecksfcs(FgNumber);
        }

        public string[] Nextstartchecksfcs(string FgNumber)
        {
            return _repo.Nextstartchecksfcs(FgNumber);
        }
    }
}
