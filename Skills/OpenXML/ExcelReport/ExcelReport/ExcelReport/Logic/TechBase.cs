using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReport.Logic
{
    class TechBase
    {
        public string techBasePath; // Путь к файлу технической базы
        public DataTable techBaseTable = new DataTable(); // Cюда записана техническая база
        public bool Optimize()
        {
            try // код помещён в блок try, потому что возможно исключение типа OpenXmlPackageException - Invalid Hyperlink
            {
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(techBasePath, true))
                {
                    techBaseTable = DataTableFromExcel.GetDataTableFromSpreadSheet(spreadSheetDocument);
                }
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink")) // если исключение сожержит "Invalid Hyperlink"
                {
                    using (FileStream fs = new FileStream(techBasePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixer.FixInvalidUri(fs, brokenUri => UriFixer.FixUri(brokenUri)); // чиним Uri
                    }
                    using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(techBasePath, true))
                    {
                        techBaseTable = DataTableFromExcel.GetDataTableFromSpreadSheet(spreadSheetDocument);
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
           
            foreach (DataRow row in techBaseTable.Rows) // Оптимизация тех. базы - убираем цифры
            {
                string product = System.Convert.ToString(row[1]); // Название продукта
                string productWithoutDigits = "";
                for (int i = 0; i < product.Length; i++) // Убираем числа
                {
                    if (!char.IsDigit(product[i]))
                    {
                        productWithoutDigits += product[i];
                    }
                }
                row[1] = productWithoutDigits;
            }
            return true;
        }

    }
}
