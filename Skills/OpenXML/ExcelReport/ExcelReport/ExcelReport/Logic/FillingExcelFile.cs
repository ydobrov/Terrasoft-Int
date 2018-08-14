using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using ExcelReport.Datastructures;

namespace ExcelReport.Logic
{
    static class FillingExcelFile
    {
        private static SLStyle styleCenter, styleHelpingToStyleCenter, styleRegular;
        // Параметры метода:
        // i - номер строки
        // sl - документ, который надо заполнять
        // dictionary - словарь с данными количества вхождений в зависимости от вендора или направления
        // columnName - подпись колонки, которую надо указать при выводе: "вендор", "направление"
        // tableName - название таблицы: имя листа+имя файла, надо которыми проводился анализ. Выводится, если пользователь выбрал опцию "по каждому месяцу отдельно"
        static public void Fill(ref int i, SLDocument sl, Dictionary<string, int> dictionary, string columnName, string tableName = null)
        {
            // Определение стилей
            
            styleCenter = sl.CreateStyle();
             styleRegular = sl.CreateStyle();
            styleCenter.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            styleCenter.Border.RightBorder.BorderStyle = styleCenter.Border.LeftBorder.BorderStyle =
                styleCenter.Border.BottomBorder.BorderStyle =
                styleCenter.Border.TopBorder.BorderStyle = BorderStyleValues.Double;
           // styleHelpingToStyleCenter.Border.LeftBorder.BorderStyle = BorderStyleValues.Double;
            styleRegular.Border.BottomBorder.BorderStyle = styleRegular.Border.TopBorder.BorderStyle =
                styleRegular.Border.RightBorder.BorderStyle = styleRegular.Border.LeftBorder.BorderStyle =
                BorderStyleValues.Thin;

            if (tableName != null)
            {
                sl.SetCellValue(i, 1, tableName);
                sl.MergeWorksheetCells(String.Format("A" + i), String.Format("B" + i));
                sl.SetCellStyle(i, 1, styleCenter);
                sl.SetCellStyle(i, 2, styleCenter);
              //  sl.SetCellStyle(i, 3, styleHelpingToStyleCenter);
                i++; // отступ - увеличиваем номер строки
            }
            sl.SetCellValue(i, 1, "Список запрашиваемых продуктов");
            

            sl.MergeWorksheetCells(String.Format("A" + i), String.Format("B" + i)); // Соединение ячеек
            sl.SetCellStyle(i, 1, styleCenter);
            sl.SetCellStyle(i, 2, styleCenter);
            //  sl.SetCellStyle(i, 3, styleHelpingToStyleCenter);
            i++;
            sl.SetCellValue(i, 1, columnName);
            sl.SetCellStyle(i, 1, styleCenter);
            sl.SetCellValue(i, 2, "Количество упоминаний");
            sl.SetCellStyle(i, 2, styleCenter);
            i++;
            DictionaryOptimization.DiscardAndOrder(ref dictionary); // Оптимизируем словарь
            foreach (KeyValuePair<string, int> vendor in dictionary)
            {
                sl.SetCellValue(i, 1, vendor.Key);
                sl.SetCellValue(i, 2, vendor.Value);
                sl.SetCellStyle(i, 1, styleRegular);
                sl.SetCellStyle(i, 2, styleRegular);
                i++; // отступ - увеличиваем номер строки
            }
            i += 2; // отступ - увеличиваем номер строки
        }

        // Параметры метода:
        // i - номер строки
        // sl - документ, который надо заполнять
        // lostProducts - ненайденные продукты (при поиске вендоров или направлений)
        // columnName - подпись колонки, которую надо указать при выводе: "вендор", "направление"
        static public void FillLostProducts(ref int i, SLDocument sl, List<LostProduct> lostProducts, string columnName)
        {
            if (lostProducts != null && lostProducts.Count != 0)
            {
                styleRegular.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Yellow, System.Drawing.Color.Yellow);
                sl.SetCellValue(i, 1, String.Format("Список ненайденных продуктов ({0})", columnName));
                sl.MergeWorksheetCells(String.Format("A" + i), String.Format("C" + i));
                sl.SetCellStyle(i, 1, styleCenter);
                sl.SetCellStyle(i, 2, styleCenter);
                sl.SetCellStyle(i, 3, styleCenter);
              //  sl.SetCellStyle(i, 4, styleHelpingToStyleCenter);
                i++;
                sl.SetCellValue(i, 1, "Название продукта");
                sl.SetCellStyle(i, 1, styleCenter);
                sl.SetCellValue(i, 2, "Название листа");
                sl.SetCellStyle(i, 2, styleCenter);
                sl.SetCellValue(i, 3, "Имя файла");
                sl.SetCellStyle(i, 3, styleCenter);
                i++; // отступ - увеличиваем номер строки

                foreach (var str in lostProducts)
                {
                    sl.SetCellValue(i, 1, str.ProductName);
                    sl.SetCellValue(i, 2, str.SheetName);
                    sl.SetCellValue(i, 3, System.IO.Path.GetFileName(str.FileName));

                    sl.SetCellStyle(i, 1, styleRegular);
                    sl.SetCellStyle(i, 2, styleRegular);
                    sl.SetCellStyle(i, 3, styleRegular);
                    i++;
                }

                i += 2; // отступ - увеличиваем номер строки

            }
            sl.AutoFitColumn(1, 3, 50); // задаём автоширину
        }
    }
}
