using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReport.Datastructures
{
    class ProductsOfInterest
    {
        public string FilePath { get; set; }
        public string SheetName { get; set; }
        public List<string> Products { get; set; }
        public Dictionary<string, int> vendors; // Вендор, количество вхождений
        public Dictionary<string, int> areas; // Направление, количество вхождений
    }
}