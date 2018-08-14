using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using SpreadsheetLight;
using ExcelReport.Logic;
using ExcelReport.Datastructures;
using System.Windows.Media;

namespace ExcelReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        TechBase techBase = new TechBase();
        Report report = new Report();
        List<ExcelFile> files; // Файлы для анализа
        List<ProductsOfInterest> products; // Список интересующих клиента продуктов в зависимости от файла и листа
        Brush red = (Brush)(new System.Windows.Media.BrushConverter().ConvertFromString("#ff0000"));
        Brush green = (Brush)(new System.Windows.Media.BrushConverter().ConvertFromString("#00FF00"));

        private void btnCreateReport_Click(object sender, RoutedEventArgs e)
        {
            if (!InputValidation())// Проверка на то, что пользователь всё выбрал и всё ввёл
            {
                return;
            }; 
            List<string> sheets_names = new List<string>();
            foreach (var item in listBox_sheets.SelectedItems)
            {
               sheets_names.Add(item.ToString()); // листы, выбранные пользователем
            }
            if (!techBase.Optimize())
            {
                OutputLabel.Content = "Ошибка! Удостоверьтесь, что файл технической базы закрыт на вашем компьютере.";
                OutputLabel.Foreground = red;
                return;
            }
            if (checkBox.IsChecked == true) // Если пользователь выбрал вывод по вендорам 
            {
                report.VendorsFilling(techBase); // Заполнение словаря Vendors
            }
            if (checkBox1.IsChecked == true) // Если пользователь выбрал вывод по направлениям 
            {
                report.AreasFilling(techBase); // Заполнение словаря Areas
            }

            string result = FillingProducts.Fill(ref products, sheets_names, files, report);
            if (result != null)
            {
                OutputLabel.Content = String.Format("Ошибка! Удостоверьтесь, что файл {0} закрыт на вашем компьютере.", System.IO.Path.GetFileName(result));
                OutputLabel.Foreground = red;
                return;
            }
            if (checkBox.IsChecked == true) // Если пользователь выбрал вывод по вендорам 
            {
                report.ComparingVendors(products, techBase.techBaseTable);
            }
            if (checkBox1.IsChecked == true) // Если пользователь выбрал вывод по направлениям 
            {
                report.ComparingAreas(products, techBase.techBaseTable);
            }
            if (!FileDialog())
            {
                OutputLabel.Content = "Отказ от сохранения файла.";
                OutputLabel.Foreground = red;
                return;
            }
            if (SLExcel())
            {
                OutputLabel.Content = "Отчет успешно сохранён!";
                OutputLabel.Foreground = green;
            }
            else
            {
                OutputLabel.Content = "Ошибка! Удостоверьтесь, что файл с таким названием закрыт на вашем компьютере.";
                OutputLabel.Foreground = red;
            }
        }

        private bool InputValidation()
        {
            if (techBase.techBasePath == null)
            {
                OutputLabel.Content = "Ошибка. Пожалуйста, удостоверьтесь, что вы загрузили файл технической базы.";
                OutputLabel.Foreground = red;
                return false;
            }
            if ((files == null) || (files.Count == 0))
            {
                OutputLabel.Content = "Ошибка. Пожалуйста, удостоверьтесь, что вы загрузили файл для анализа.";
                OutputLabel.Foreground = red;
                return false;
            }
            if (listBox_sheets.SelectedItems.Count == 0)
            {
                OutputLabel.Content = "Ошибка. Пожалуйста, удостоверьтесь, что вы выбрали листы для анализа.";
                OutputLabel.Foreground = red;
                return false;
            }
            if (!(((radioButton.IsChecked == true) || (radioButton_Copy.IsChecked == true)) &&
                     ((checkBox.IsChecked == true) || (checkBox1.IsChecked == true))))
            {
                OutputLabel.Content = "Неправильный ввод! Пожалуйста, удостоверьтесь, что вы проставили все флаги!";
                OutputLabel.Foreground = red;
                return false;
            }
            return true;
        }

        private bool FileDialog()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "Report"; // Default file name
            saveFileDialog.DefaultExt = ".xlsx"; // Default file extension
            saveFileDialog.Filter = "XML-File | *.xlsx";
            Nullable<bool> result = saveFileDialog.ShowDialog();
            // Process save file dialog box results
            if (result == true)
            {
                report.DestinationPathReport = saveFileDialog.FileName; // задаём путь для файла с отчётом
                return true;
            }
            return false;
        }

        private void btnOpenFiles_Click(object sender, RoutedEventArgs e)
        {
            OutputLabel.Content = "";
            OpenFileDialog openFileDialog = new OpenFileDialog(); // Открывается диалоговое окно
            openFileDialog.Multiselect = true; // возможность выбрать несколько файлов
            openFileDialog.Filter = "Excel Files|*.xlsx;"; // возможность выбрать только xlsx файлы
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // Директория по умолчанию
            if (openFileDialog.ShowDialog() == true)
            {
                files = new List<ExcelFile>();
                foreach (string filename in openFileDialog.FileNames)
                {
                    List<string> sheet_names = SheetNames.GetSheets(filename); // узнаём наименования листов
                    if (sheet_names == null)
                    {
                        OutputLabel.Content = String.Format("Ошибка! Удостоверьтесь, что файл {0} закрыт на вашем компьютере.", System.IO.Path.GetFileName(filename));
                        OutputLabel.Foreground = red;
                        continue;
                    }

                    lbFiles.Items.Add(System.IO.Path.GetFileName(filename)); // Добавляем на элемент lbFiles названия файлов

                    List<string> sheets = new List<string>();
                    foreach (string sheet in sheet_names)
                    {
                        listBox_sheets.Items.Add(sheet); // listBox_sheets - элемент в xaml-коде, добавляем туда названия листов
                        sheets.Add(sheet);
                    }
                    files.Add(new ExcelFile() { FilePath = filename, Sheets = sheets });
                }
            }
        }

        private void btnOpenTechBase_Click(object sender, RoutedEventArgs e) // Обработка нажатия по кнопке для открытия файла технической базы
        {
            OutputLabel.Content = "";
            OpenFileDialog openFileDialog = new OpenFileDialog(); // открываем диалоговое окно для выбора файлов
            openFileDialog.Multiselect = false; // нельзя выбрать больше одного файла
            openFileDialog.Filter = "Excel Files|*.xlsx;"; // можно выбрать только xlsx файлы
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // директория для открытия по умолчанию 
            if (openFileDialog.ShowDialog() == true)
            {
                techBase.techBasePath = openFileDialog.FileName; // путь к файлу находится в переменной techBasePath
                lbTechBase.Items.Clear();
                lbTechBase.Items.Add(System.IO.Path.GetFileName(techBase.techBasePath)); // добавляем на элемент lbTechBase
            }
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void radioButton_Checked_1(object sender, RoutedEventArgs e)
        {
            OutputLabel.Content = "";
        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {
            OutputLabel.Content = "";
        }

        private void listBox_sheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            OutputLabel.Content = "";
        }

        public bool SLExcel()
        {
            int i = 2; // номер cтроки
            SLDocument sl = new SLDocument();

            if (radioButton.IsChecked == true) // Если пользователь указал опцию вывода по каждому месяцу (листу)
            {
                foreach (ProductsOfInterest product in products)
                {
                    string tableName = System.IO.Path.GetFileName(product.FilePath) + " - " + product.SheetName;
                    if (checkBox.IsChecked == true) // Если пользователь выбрал вывод по вендорам 
                    {
                        FillingExcelFile.Fill(ref i, sl, product.vendors, "Вендор", tableName);
                      
                    }
                    if (checkBox1.IsChecked == true) // Если пользователь выбрал вывод по направлениям 
                    {
                        FillingExcelFile.Fill(ref i, sl, product.areas, "Направление", tableName);
                    }
                }
                FillingExcelFile.FillLostProducts(ref i, sl, report.LostProductsByVendor, "Вендор");
                FillingExcelFile.FillLostProducts(ref i, sl, report.LostProductsByArea, "Направление");
            }
            else
            {
                if (checkBox.IsChecked == true) // Если пользователь выбрал вывод по вендорам 
                {
                    FillingExcelFile.Fill(ref i, sl, report.Vendors, "Вендор");
                }
                if (checkBox1.IsChecked == true) // Если пользователь выбрал вывод по направлениям 
                {
                    FillingExcelFile.Fill(ref i, sl, report.Areas, "Направление");
                }
                FillingExcelFile.FillLostProducts(ref i, sl, report.LostProductsByVendor, "Вендор");
                FillingExcelFile.FillLostProducts(ref i, sl, report.LostProductsByArea, "Направление");
            }
            try
            {
                sl.SaveAs(report.DestinationPathReport); // Сохраняем файл
            }
            catch (System.IO.IOException)
            {
                return false; // Исключение свидетельсвует о том, что файл с таким названем используется другим приложением
            }
            System.Diagnostics.Process.Start(report.DestinationPathReport); // Открываем файл
            return true;
            }


      }
}