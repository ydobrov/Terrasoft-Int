using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReport.Datastructures;
using System.Data;

namespace ExcelReport.Logic
{
    class Report
    {
        public string DestinationPathReport; // Путь для того, чтобы сохранить файл с отчетом
        public  Dictionary<string, int> Vendors; // Вендор, количество вхождений
        public  Dictionary<string, int> Areas; // Направление, количество вхождений
        public  List<LostProduct> LostProductsByVendor; // Продукты, не имеющие соответсвия с вендорами
        public  List<LostProduct> LostProductsByArea; // Продукты, не имеющие соответсвия с направлениями

        // Правильно было бы сократить шесть этих методов до трёх, так как три пары практически идентичны,
        // но такой код кажется более понятным, хоть он и выглядит объёмным

        public void VendorsFilling(TechBase techBase)
        {
            Vendors = new Dictionary<string, int>();
            List<string> vendorsList = new List<string>();
            foreach (DataRow row in techBase.techBaseTable.Rows)
            {
                if ((string)row[0] != "")
                    vendorsList.Add((string)row[0]); // В первом столбце вендоры
            }
            vendorsList = vendorsList.Distinct().ToList();
            foreach (string vendor in vendorsList)
            {
                Vendors.Add(vendor, 0);
            }
        }

        public void AreasFilling(TechBase techBase)
        {
            Areas = new Dictionary<string, int>();
            List<string> areasList = new List<string>();
            foreach (DataRow row in techBase.techBaseTable.Rows)
            {
                try // возможно исключение "нельзя привести DBNull к string"
                {
                    if ((string)row[2] != "") // В третьем столбце направления
                        areasList.Add((string)row[2]); // В третьем столбце направления
                }
                catch (Exception)
                { }
            }
            areasList = areasList.Distinct().ToList<string>(); // с помощью Distinct() отбрасываем повторяющиеся значения
            foreach (string area in areasList)
            {
                Areas.Add(area, 0);
            }
        }
        
        public void ComparingVendors(List<ProductsOfInterest> products, DataTable techBase)
        {
            LostProductsByVendor = new List<LostProduct>();
            foreach (ProductsOfInterest prod in products)
            {
                foreach (string prodName in prod.Products)
                {
                    foreach (DataRow row in techBase.Rows)
                    {
                        if ((((string)row[0]) != "") && ((((string)row[1]).ToLower().Contains(prodName.ToLower())) || (prodName.ToLower().Contains(((string)row[1]).ToLower()))
                            || (((string)row[0]).ToLower().Contains(prodName.ToLower())) || (prodName.ToLower().Contains(((string)row[0]).ToLower()))))// Сравнение названия продукта с тех.базы и названия интересующего клиента продукта
                        {
                            MatchingSearchVendors(row, prod);
                            break;
                        }
                        if (techBase.Rows[techBase.Rows.Count - 1] == row)
                        {
                            LostProductsByVendor.Add(new LostProduct()
                            {
                                FileName = prod.FilePath,
                                SheetName = prod.SheetName,
                                ProductName = prodName
                            });
                        }
                    }
                }
            }
        }

        public void MatchingSearchVendors(DataRow row, ProductsOfInterest prod)
        {
            foreach (var vendor in Vendors)
            {
                if (vendor.Key == (string)row[0])
                {
                    Vendors[vendor.Key] += 1;
                    prod.vendors[vendor.Key] += 1;
                    return;
                }
            }
            LostProductsByVendor.Add(new LostProduct()
            {
                FileName = prod.FilePath,
                SheetName = prod.SheetName,
                ProductName = (string)row[1]
            });
        }

        public void ComparingAreas(List<ProductsOfInterest> products, DataTable techBase)
        {
            LostProductsByArea = new List<LostProduct>();
            foreach (ProductsOfInterest prod in products)
            {
                foreach (string prodName in prod.Products)
                {
                    foreach (DataRow row in techBase.Rows)
                    {
                        if ((((string)row[0]) != "") && ((((string)row[1]).ToLower().Contains(prodName.ToLower())) || (prodName.ToLower().Contains(((string)row[1]).ToLower()))
                        || (((string)row[0]).ToLower().Contains(prodName.ToLower())) || (prodName.ToLower().Contains(((string)row[0]).ToLower()))))// Сравнение названия продукта с тех.базы и названия интересующего клиента продукта
                        {
                            MatchingSearchAreas(row, prod);
                            break;
                        }
                        if (techBase.Rows[techBase.Rows.Count - 1] == row)
                        {
                            LostProductsByArea.Add(new LostProduct()
                            {
                                FileName = prod.FilePath,
                                SheetName = prod.SheetName,
                                ProductName = prodName
                            });
                        }
                    }
                }
            }
        }

        public void MatchingSearchAreas(DataRow row, ProductsOfInterest prod)
        {
            foreach (var area in Areas)
            {
                if (area.Key == System.Convert.ToString(row[2]));
                {
                    Areas[area.Key] += 1;
                    prod.areas[area.Key] += 1;
                    return;
                }
            }
            LostProductsByArea.Add(new LostProduct()
            {
                FileName = prod.FilePath,
                SheetName = prod.SheetName,
                ProductName = (string)row[1]
            });
        }
    }
}
