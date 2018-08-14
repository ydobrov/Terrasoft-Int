using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReport
{
    class SheetNames
    {
        public static List<string> GetSheets(string fileName) // узнаём имена листов исходя из имени файла
        {
            try // код помещён в блок try, потому что возможно исключение типа OpenXmlPackageException - Invalid Hyperlink
            {
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, true))
                {
                    return GetSheetNames(spreadSheetDocument); // узнаём имена листов исходя из SpreadsheetDocument'a
                }
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink")) // если исключение сожержит "Invalid Hyperlink"
                {
                    using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixer.FixInvalidUri(fs, brokenUri => UriFixer.FixUri(brokenUri)); // чиним Uri
                    }
                    using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, true))
                    {
                        return GetSheetNames(spreadSheetDocument); // узнаём имена листов исходя из SpreadsheetDocument'a
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
            return null;
        }

        private static List<string> GetSheetNames(SpreadsheetDocument spreadSheetDocument) // узнаём имена листов исходя из SpreadsheetDocument'a
        {
            List<string> sheets = new List<string>();
            int sheetIndex = 0;
            foreach (WorksheetPart worksheetpart in spreadSheetDocument.WorkbookPart.WorksheetParts)
            {
                Worksheet worksheet = worksheetpart.Worksheet;

                // Grab the sheet name each time through your loop
                string sheetName = spreadSheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name;
                sheets.Add(sheetName);
                sheetIndex++;
            }
            return sheets;
        }


    }
}
