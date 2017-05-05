using System;
using System.Collections.Generic;
using System.IO;
using RL.FileGenerators.Excel;

namespace ConsoleDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateExcelFileFromCollection();
        }

        private static void CreateExcelFileFromCollection()
        {
            List<Product> products = new List<Product>(3);
            products.Add(new Product()
            {
                Code = "P_A",
                Discount = 0.1F,
                Id = 1,
                Name = "ABC",
                Price = 6000.85m,
                AvailableDate = DateTime.Now
            });
            products.Add(new Product()
            {
                Code = "P_B",
                Discount = 0.2F,
                Id = 2,
                Name = "XYZ",
                Price = 5500.50m,
                AvailableDate = DateTime.Now.AddDays(-3),
                IsOffline = true,
                NullableBool = true
            });
            products.Add(new Product()
            {
                Code = "P_C",
                Discount = 0.22F,
                Id = 3,
                Name = "123",
                Price = null,
                AvailableDate = DateTime.Now.AddDays(-3),
                IsOffline = true,
                NullableBool = true
            });
            products.Add(new Product()
            {
                Code = "P_D",
                Discount = 0.35F,
                Id = 4,
                Name = "456",
                Price = 5800.0m,
                AvailableDate = DateTime.Now.AddDays(-5)
            });

            string fileName = Path.Combine(Directory.GetCurrentDirectory(), "Test.xlsx");

            // create .xlsx file as MemoryStream
            using (MemoryStream ms = ExcelFileGenerator.CreateStream(products, "Products"))
            {
                using (FileStream fs = File.Create(fileName))
                {
                    ms.WriteTo(fs);
                }
            }

            // create .xlsx file
            //ExcelGenerator.CreateFile(products, "Products", fileName);
        }
    }
}