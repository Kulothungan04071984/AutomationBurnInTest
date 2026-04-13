using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BurnInTestValidate
{
    public class services
    {
        public string stageid { get; set; }
        public string StageName { get; set; }
        public string CurrentStageid { get; set; }
        public string CurrentStageName { get; set; }

        public Dictionary<bool, int> resultset { get; set; }
    }
}
