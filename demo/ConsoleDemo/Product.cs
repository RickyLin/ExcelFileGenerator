using System;
using RL.FileGenerators.Excel;

namespace ConsoleDemo
{
    public class Product
    {
        [ExcelColumn("Product Id", 4)]
        public int Id { get; set; }

        public string Code { get; set; }

        [ExcelColumn("Product Name", 1)]
        public string Name { get; set; }

        [ExcelColumn("Product Price", 3, NumberingFormatString = "#,##0.00")]
        public decimal? Price { get; set; }

        [ExcelColumn("Product Discount", 2)]
        public float Discount { get; set; }

        [ExcelColumn("Available Date", 5, NumberFormatId = 22)]
        public DateTime AvailableDate { get; set; }

        [ExcelColumn("Is Offline", 6)]
        public bool IsOffline { get; set; }

        [ExcelColumn("NullableBool", 7)]
        public bool? NullableBool { get; set; }
    }
}
