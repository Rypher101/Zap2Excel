using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zap2Excel
{
    public class ScanTable
    {
        public string Severity { get; set; }
        public string Vulnerability { get; set; }
        public string Description { get; set; }

        public List<string[]> URLs = new List<string[]>();
        public int CWE { get; set; }
    }
}
