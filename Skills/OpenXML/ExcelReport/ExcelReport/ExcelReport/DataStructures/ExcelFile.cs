using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReport.Datastructures
{
    class ExcelFile
    {
        public string FilePath { get; set; }
        public List<string> Sheets { get; set; }
    }
}
