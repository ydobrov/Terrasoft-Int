using DocumentFormat.OpenXml.Packaging;
using ExcelReport.Datastructures;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReport.Logic
{
    static class FillingProducts
    {
       static string[] stringSeparators = new string[] { ", " }; // разделитель строк, содержащих интересуюие продукты

        public static string Fill(ref List<ProductsOfInterest> products, List<string> sheets_names, List<ExcelFile> files, Report report)
        {
            products = new List<ProductsOfInterest>();
            foreach (string sheetFromListBox in sheets_names)
            {
                foreach (ExcelFile file in files)
                {
                    foreach (string sheetFromFile in file.Sheets)
                    {
                        if (sheetFromListBox.Equals(sheetFromFile))
                        {
                            try // код помещён в блок try, потому что возможно исключение типа OpenXmlPackageException - Invalid Hyperlink
                            {
                                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(file.FilePath, true))
                                {
                                    HelpFill(spreadSheetDocument, sheetFromFile, file, ref products, report);
                                }
                            }
                            catch (OpenXmlPackageException e)
                            {
                                if (e.ToString().Contains("Invalid Hyperlink")) // если исключение сожержит "Invalid Hyperlink"
                                {
                                    using (FileStream fs = new FileStream(file.FilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                                    {
                                        UriFixer.FixInvalidUri(fs, brokenUri => UriFixer.FixUri(brokenUri)); // чиним Uri
                                    }
                                    using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(file.FilePath, true))
                                    {
                                        HelpFill(spreadSheetDocument, sheetFromFile, file, ref products, report);
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                return file.FilePath;
                            }
                        }
                    }
                }
            }
            return null;
        }

        static public void HelpFill(SpreadsheetDocument spreadSheetDocument, string sheetFromFile, ExcelFile file, ref List<ProductsOfInterest> products, Report report)
        {
            DataTable dt = DataTableFromExcel.GetDataTableFromSpreadSheet(spreadSheetDocument, sheetName: sheetFromFile);
            List<string> productsFromSheet = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                try // возможно исключение "нельзя привести DBNull к string"
                {
                    if ((string)row[2] != "") // В третьем столбце интересующие продукты
                    {
                        if ((string)row[2] == "-") // Клиент не интересовался продуктами
                            continue;
                        string product = (string)row[2];
                        string productWithoutDigits = "";
                        for (int i = 0; i < product.Length; i++)
                        {
                            if (!char.IsDigit(product[i]))
                            {
                                productWithoutDigits += product[i];
                            }
                        }
                        string[] productWithoutDigitsSeparated = productWithoutDigits.Split(stringSeparators, StringSplitOptions.None);
                        foreach (string prod in productWithoutDigitsSeparated)
                            productsFromSheet.Add(prod);
                    }
                }
                catch (Exception e)
                { }
            }

            products.Add(new ProductsOfInterest()
            {
                FilePath = file.FilePath,
                SheetName = sheetFromFile,
                Products = productsFromSheet,
                vendors = (report.Vendors != null) ? new Dictionary<string, int>(report.Vendors) : new Dictionary<string, int>(),
                areas = (report.Areas != null) ? new Dictionary<string, int>(report.Areas) : new Dictionary<string, int>()
            });
        }
    }
}
