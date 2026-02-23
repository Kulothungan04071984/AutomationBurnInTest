using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BurnInTestValidate
{
    public class PassmarkHistory
    {
        public string FgNumber { get; set; }
        public string CustomerSerialNumber { get; set; }
        public string PCBAID { get; set; }
        public string DiskPartition { get; set; }
        public string CrystalReport { get; set; }
        public string read_one { get; set; }
        public string read_two { get; set; }
        public string read_three { get; set; }
        public string read_four { get; set; }
        public string write_one { get; set; }
        public string write_two { get; set; }
        public string write_three { get; set; }
        public string write_four { get; set; }
        public string burnintest { get; set; }
        public string overall_result { get; set; }
        public string CreatedBy { get; set; }
    }
}
