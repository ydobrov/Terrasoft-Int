using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Text.RegularExpressions;
namespace ExcelReport
{

    /// <summary>
    /// Source: https://www.experts-exchange.com/questions/28732660/Reading-Excel-Spreadsheet-using-DocumentFormat-OpenXML-skips-columns-that-don%27t-have-data-in-them.html#answer41016147-20
    /// </summary>

    static class DataTableFromExcel
    {
        /// <summary>Gets the data table from spread sheet.</summary>
        /// <param name="document">The document.</param>
        /// <param name="columns">The columns.</param>
        /// <param name="excludeHeader">if set to <c>true</c> [exclude header].</param>
        /// <returns>DataTable.</returns>
        public static DataTable GetDataTableFromSpreadSheet(this SpreadsheetDocument document, int columns = -1, bool excludeHeader = true, string sheetName = "")
        {
            var results = new DataTable();
            try
            {
                var sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string id;
                if (sheetName != "")
                {
                    id = document.WorkbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
                }
                else
                {
                    id = sheets.First().Id.Value;
                }
                var part = (WorksheetPart)document.WorkbookPart.GetPartById(id);
                var sheet = part.Worksheet;
                var data = sheet.GetFirstChild<SheetData>();
                var rows = data.Descendants<Row>();
                if (rows.Count() != 0)
                {
                    var colCount = rows.First().Cast<Cell>().Count();
                    if (columns > colCount || columns <= 0)
                        columns = colCount;

                    foreach (var cell in rows.First().Cast<Cell>().Take(columns))
                        results.Columns.Add(cell.GetValue(document));

                    foreach (var row in rows.Skip(Convert.ToInt32(excludeHeader)))
                        results.Rows.Add((from cell in row.GetCells().Take(columns) select cell.GetValue(document)).ToArray());
                }
            }
            catch (Exception)
            {
                results = new DataTable();
            }
            return results;
        }

        /// <summary>Gets the value.</summary>
        /// <param name="cell">The cell.</param>
        /// <param name="document">The document.</param>
        /// <returns>System.String.</returns>

        public static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }

        public static string GetValue(this Cell cell, SpreadsheetDocument document)
        {
            string result = string.Empty;
            try
            {
                if (cell != null && cell.ChildElements.Count != 0)
                {
                    var part = document.WorkbookPart.SharedStringTablePart;
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        result = part.SharedStringTable.ChildElements[Int32.Parse(cell.CellValue.InnerText)].InnerText;
                    else
                        result = cell.CellValue.InnerText;
                }
            }
            catch (Exception)
            {
                result = string.Empty;
            }
            return result;
        }

        /// <summary>Gets the cells from a row.</summary>
        /// <param name="row">The row.</param>
        /// <returns>IEnumerable&lt;Cell&gt;.</returns>
        public static IEnumerable<Cell> GetCells(this Row row)
        {
            int count = 0;
            foreach (Cell cell in row.Descendants<Cell>())
            {
                int index = cell.CellReference.GetColumnName().ConvertColumnNameToNumber();
                for (; count < index; count++)
                {
                    // Null or empty cell reference encountered, replace with a new cell that contains a default value.
                    yield return new Cell() { DataType = null, CellValue = new CellValue(string.Empty) };
                }
                yield return cell;
                count++;
            }
        }

        /// <summary>Gets the name of the column.</summary>
        /// <param name="reference">The cell reference.</param>
        /// <returns>System.String.</returns>
        private static string GetColumnName(this StringValue reference)
        {
            return new Regex("(?i)[a-z]+").Match(reference).Value;
        }

        /// <summary>Converts the column name to number.</summary>
        /// <param name="name">The name.</param>
        /// <returns>System.Int32.</returns>
        /// <exception cref="System.ArgumentException"></exception>
        private static int ConvertColumnNameToNumber(this string name)
        {
            var expression = new Regex("^[A-Z]+$");
            if (!expression.IsMatch(name))
                throw new ArgumentException();

            char[] letters = name.ToCharArray();
            Array.Reverse(letters);

            int result = 0;
            for (int i = 0; i < letters.Length; i++)
                result += ((i == 0) ? letters[i] - 65 : letters[i] - 64) * (int)Math.Pow(26, i);
            return result;
        }
    }
}