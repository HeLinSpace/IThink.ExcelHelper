﻿using System;
using System.Collections.Generic;
using System.Linq;
using H.Npoi.ExcelHelper;
using NPOI.XSSF.UserModel;

namespace IThink.ExcelHelper.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var template = "./AppData/TestImport.xlsx";

            // test import
            using (var w = NExcelHelper.Open(template))
            {
                var res = w.Import<TestImport>(0, 1, 4, RowCheck, ImportBussiness);
            }

            // test export
            var list = new List<TestExport>();

            #region get data

            list.Add(new TestExport
            {
                Age = 13,
                Birthday = DateTime.Now,
                Name = "Nancy",
                Sex = 1
            });

            list.Add(new TestExport
            {
                Age = 12,
                Birthday = DateTime.Now,
                Name = "Tom",
                Sex = 0
            });

            #endregion

            var workbook = list.ExportWorkbook(template, 1, (workbook, index) =>
             {
                 var cellStyle = NExcelHelper.CreateCellStyle(workbook);
                 if (index == 3)
                 {
                     var dataFormatCustom = (XSSFDataFormat)workbook.CreateDataFormat();
                     cellStyle.DataFormat = dataFormatCustom.GetFormat(NExcelHelper.CellFormat.DefaultShortDate);
                 }

                 return cellStyle;
             });

            workbook.Save("./AppData/TestExport.xlsx");
        }

        private static void RowCheck(TestImport item, ImportSheetInfo excelImport)
        {
            // do some  check like this 
            if (item.Birthday == null)
            {
                item.ErrorMsg = "the birthday can not be null.";
            }
        }

        private static dynamic ImportBussiness(List<TestImport> list, ImportSheetInfo excelImport)
        {
            // has error
            if (list.Any(s => !string.IsNullOrEmpty(s.ErrorMsg)))
            {
                // write the error col and save as file
                excelImport.WriteErrorFile(list, "./import/error.xlsx");

                return false;
            }

            // do bussiness

            return true;
        }
    }

    internal class TestExport : IExportModel
    {
        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(0)]
        public string Name { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(1)]
        public int Age { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(2)]
        public int Sex { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(3)]
        public DateTime? Birthday { get; set; }
    }

    internal class TestImport : ImportBaseModel
    {
        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(0)]
        public string Name { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(1)]
        public int Age { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(2)]
        public int Sex { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [ColumnProperty(3)]
        public DateTime? Birthday { get; set; }
    }
}
