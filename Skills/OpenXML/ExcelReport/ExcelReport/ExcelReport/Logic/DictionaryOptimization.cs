using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReport.Logic
{
    static class DictionaryOptimization
    {
        // Откидываем пары с нулевыми значениями и сортируем
        static public void DiscardAndOrder(ref Dictionary<string, int> dict)
        {
            Dictionary<string, int> dictionary = new Dictionary<string, int>();
            foreach (KeyValuePair<string, int> d in dict)
            {
                if (d.Value > 0)
                {
                    dictionary.Add(d.Key, d.Value);
                }
            }
            var ordered = dictionary.OrderByDescending(x => x.Value);
            dict = ordered.ToDictionary(t => t.Key, t => t.Value);
        }

    }
}
