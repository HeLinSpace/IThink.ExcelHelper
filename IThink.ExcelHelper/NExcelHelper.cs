using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.AspNetCore.Http;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace H.Npoi.ExcelHelper
{
    /// <summary>
    /// Excel Helper for NPOI.
    /// </summary>
    public static class NExcelHelper
    {
        /// <summary>
        /// open workbook by full path
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static ExcelImport Open(string path)
        {
            return new ExcelImport(path);
        }

        /// <summary>
        /// open workbook by request
        /// </summary>
        /// <param name="file">request formfile</param>
        /// <returns></returns>
        public static ExcelImport Open(IFormFile file)
        {
            return new ExcelImport(file);
        }

        /// <summary>
        /// open workbook by stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="isXlsx"></param>
        /// <returns></returns>
        public static ExcelImport Open(Stream stream, bool isXlsx = true)
        {
            return new ExcelImport(stream, isXlsx);
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="filePath"></param>
        public static void Save(this IWorkbook workbook, string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="workbook"></param>
        public static byte[] Save(this IWorkbook workbook)
        {
            using (var fs = new MemoryStream())
            {
                workbook.Write(fs);

                return fs.ToArray();
            }
        }

        /// <summary>
        /// 执行导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="dataRowIndex">数据行</param>
        /// <param name="templateFullPath">模板全路径</param>
        /// <param name="cellStyleFunc">T1:current workbook  T2:col index</param>
        /// <returns></returns>
        public static byte[] Export<T>(this List<T> dataList, string templateFullPath, int dataRowIndex, Func<IWorkbook, int, ICellStyle> cellStyleFunc) where T : IExportModel
        {
            var exportData = dataList.GetExportData(dataRowIndex);

            using (var templatefs = new FileStream(templateFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var xssfWorkbook = new XSSFWorkbook(templatefs);

                var workSheet = xssfWorkbook.GetSheetAt(0);

                foreach (var rowData in exportData)
                {
                    var row = workSheet.CreateRow(rowData.RowNo);
                    var cols = rowData.Cols.OrderBy(s => s.ColNo);
                    foreach (var colData in cols)
                    {
                        var cell = row.CreateCell(colData.ColNo);
                        SetValue(cell, colData.Value);

                        if (cellStyleFunc != null)
                        {
                            var cellstyle = cellStyleFunc(xssfWorkbook, colData.ColNo);
                            cell.CellStyle = cellstyle;
                        }
                    }
                }

                using (var memoryStream = new MemoryStream())
                {
                    xssfWorkbook.Write(memoryStream);
                    xssfWorkbook.Close();

                    return memoryStream.ToArray();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="templateFullPath"></param>
        /// <param name="dataRowIndex"></param>
        /// <param name="cellStyleFunc">T1:current workbook T2:col index</param>
        /// <returns></returns>
        public static XSSFWorkbook ExportWorkbook<T>(this List<T> dataList, string templateFullPath, int dataRowIndex, Func<IWorkbook, int, ICellStyle> cellStyleFunc) where T : IExportModel
        {
            var exportData = dataList.GetExportData(dataRowIndex);

            using (var templatefs = new FileStream(templateFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var xssfWorkbook = new XSSFWorkbook(templatefs);

                var workSheet = xssfWorkbook.GetSheetAt(0);

                foreach (var rowData in exportData)
                {
                    var row = workSheet.CreateRow(rowData.RowNo);
                    var cols = rowData.Cols.OrderBy(s => s.ColNo);
                    foreach (var colData in cols)
                    {
                        var cell = row.CreateCell(colData.ColNo);
                        SetValue(cell, colData.Value);

                        if (cellStyleFunc != null)
                        {
                            var cellstyle = cellStyleFunc(xssfWorkbook, colData.ColNo);
                            cell.CellStyle = cellstyle;
                        }
                    }
                }

                return xssfWorkbook;
            }
        }

        /// <summary>
        /// 执行导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="dataRowIndex">数据行</param>
        /// <param name="templateFullPath">模板全路径</param>
        /// <param name="cellStyleFunc">单元格格式</param>
        /// <returns></returns>
        public static XSSFWorkbook ExportWorkbook<T>(this List<T> dataList, string templateFullPath, int dataRowIndex, Func<IWorkbook, ICellStyle> cellStyleFunc) where T : IExportModel
        {
            var exportData = dataList.GetExportData(dataRowIndex);

            using (var templatefs = new FileStream(templateFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var xssfWorkbook = new XSSFWorkbook(templatefs);

                var workSheet = xssfWorkbook.GetSheetAt(0);

                var cellstyle = CreateCellStyle(xssfWorkbook);

                if (cellStyleFunc != null)
                {
                    cellstyle = cellStyleFunc(xssfWorkbook);
                }

                foreach (var rowData in exportData)
                {
                    var row = workSheet.CreateRow(rowData.RowNo);
                    var cols = rowData.Cols.OrderBy(s => s.ColNo);
                    foreach (var colData in cols)
                    {
                        var cell = row.CreateCell(colData.ColNo);
                        SetValue(cell, colData.Value);
                        cell.CellStyle = cellstyle;
                    }
                }

                return xssfWorkbook;
            }
        }

        /// <summary>
        /// 执行导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="dataRowIndex">数据行</param>
        /// <param name="templateFullPath">模板全路径</param>
        /// <param name="cellStyle">单元格格式</param>
        /// <param name="sheetNo">work sheet</param>
        /// <returns></returns>
        public static byte[] Export<T>(this List<T> dataList, string templateFullPath, int dataRowIndex, ICellStyle cellStyle = null, int sheetNo = 0) where T : IExportModel
        {
            var exportData = dataList.GetExportData(dataRowIndex);

            using (var templatefs = new FileStream(templateFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var xssfWorkbook = new XSSFWorkbook(templatefs);

                var workSheet = xssfWorkbook.GetSheetAt(sheetNo);

                if (cellStyle == null)
                {
                    cellStyle = CreateCellStyle(xssfWorkbook);
                }

                foreach (var rowData in exportData)
                {
                    var row = workSheet.CreateRow(rowData.RowNo);
                    var cols = rowData.Cols.OrderBy(s => s.ColNo);
                    foreach (var colData in cols)
                    {
                        var cell = row.CreateCell(colData.ColNo);
                        SetValue(cell, colData.Value);

                        cell.CellStyle = cellStyle;
                    }
                }

                using (var memoryStream = new MemoryStream())
                {
                    xssfWorkbook.Write(memoryStream);
                    xssfWorkbook.Close();

                    return memoryStream.ToArray();
                }
            }
        }

        /// <summary>
        /// 执行导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="dataRowIndex">数据行</param>
        /// <param name="templateFullPath">模板全路径</param>
        /// <param name="cellStyle">单元格格式</param>
        /// <param name="sheetNo">work sheet</param>
        /// <returns></returns>
        public static XSSFWorkbook ExportWorkbook<T>(this List<T> dataList, string templateFullPath, int dataRowIndex, ICellStyle cellStyle = null, int sheetNo = 0) where T : IExportModel
        {
            var exportData = dataList.GetExportData(dataRowIndex);

            using (var templatefs = new FileStream(templateFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var xssfWorkbook = new XSSFWorkbook(templatefs);

                var workSheet = xssfWorkbook.GetSheetAt(sheetNo);

                if (cellStyle == null)
                {
                    cellStyle = CreateCellStyle(xssfWorkbook);
                }

                foreach (var rowData in exportData)
                {
                    var row = workSheet.CreateRow(rowData.RowNo);
                    var cols = rowData.Cols.OrderBy(s => s.ColNo);
                    foreach (var colData in cols)
                    {
                        var cell = row.CreateCell(colData.ColNo);
                        SetValue(cell, colData.Value);

                        cell.CellStyle = cellStyle;
                    }
                }

                return xssfWorkbook;
            }
        }

        #region Tools

        /// <summary>
        /// 获取NPOI的单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static dynamic GetValue(this ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (DateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue;
                    }
                    else
                    {
                        return cell.NumericCellValue;
                    }
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);

                        if (cell.CellType == CellType.Error)
                        {
                            return null;
                        }
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.StringCellValue.ToString();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型
                    return cell.ToString();
            }
        }

        /// <summary>
        /// set the cell value.auto match the value type.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param> 
        public static void SetValue(this ICell cell, dynamic value)
        {
            try
            {
                if (value == null)
                {
                    cell.SetCellValue("");
                }
                else
                {
                    if (value.GetType() == typeof(string))
                    {
                        var strVal = value.ToString();
                        if (strVal.StartsWith("="))
                        {
                            cell.SetCellFormula(strVal.TrimStart('='));
                        }
                        else
                        {
                            cell.SetCellValue(value.ToString());
                        }
                    }
                    else if (value.GetType() == typeof(int))
                    {
                        cell.SetCellValue(Convert.ToInt32(value));
                    }
                    else if (value.GetType() == typeof(float))
                    {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else if (value.GetType() == typeof(double))
                    {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else if (value.GetType() == typeof(decimal))
                    {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else if (value.GetType() == typeof(bool))
                    {
                        cell.SetCellValue(Convert.ToBoolean(value));
                    }
                    else if (value.GetType() == typeof(DateTime))
                    {
                        cell.SetCellValue(Convert.ToDateTime(value));
                    }
                    else
                    {
                        cell.SetCellValue(Convert.ToString(value));
                    }
                }
            }
            catch
            {
                return;
            }
        }

        /// <summary>
        /// create the CellStyle by Workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="size">default 10</param>
        /// <param name="isBold"></param>
        /// <param name="isBorder"></param>
        /// <param name="isWrapText"></param>
        /// <param name="backColor"></param>
        /// <param name="cellFormat"></param>
        /// <param name="verticalAlignment"></param>
        /// <param name="horizontalAlignment"></param>
        /// <param name="fontName"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public static ICellStyle CreateCellStyle(this IWorkbook workbook, int size = 10, bool isBold = false, bool isBorder = true, bool isWrapText = true, IColor backColor = null, string cellFormat = "@", VerticalAlignment verticalAlignment = VerticalAlignment.Center, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left, string fontName = "宋体", IColor fontColor = null)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException("Workbook");
            }

            ICellStyle style = workbook.CreateCellStyle();

            //添加表格线
            if (isBorder)
            {
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
                style.BorderTop = BorderStyle.Thin;
            }

            IFont font = workbook.CreateFont();
            //设置字体颜色
            if (fontColor != null)
            {
                font.Color = ((XSSFColor)fontColor).Index;
            }
            else
            {
                font.Color = new XSSFColor(Color.Black).Index;
            }

            //设置字体粗细
            font.IsBold = isBold;

            //设置字体大小
            font.FontHeightInPoints = size;
            font.FontName = fontName;
            style.SetFont(font);

            //设置背景颜色
            if (backColor != null)
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = ((XSSFColor)backColor).Index;
                ((XSSFColor)style.FillForegroundColorColor).SetRgb(((XSSFColor)backColor).RGB);
            }

            //设置值是否换行
            style.WrapText = isWrapText;

            //设置垂直位置
            style.VerticalAlignment = verticalAlignment;
            //设置水平位置
            style.Alignment = horizontalAlignment;

            //设置值格式
            XSSFDataFormat dataFormatCustom = (XSSFDataFormat)workbook.CreateDataFormat();
            style.DataFormat = dataFormatCustom.GetFormat(cellFormat);

            return style;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        public static decimal? ToDecimal(dynamic v)
        {
            if (v == null)
            {
                return null;
            }

            return Convert.ToDecimal(v);
        }

        /// <summary>
        /// CellFormat
        /// </summary>
        public struct CellFormat
        {
            /// <summary>
            /// 1000 -> 1,000
            /// </summary>
            public const string Num_Thou = "#,##0_ ";

            /// <summary>
            /// 1000 -> 1,000.0
            /// </summary>
            public const string Num_ThouDec = "#,##0.0_ ";

            /// <summary>
            /// 1000 -> 1,000.0.
            /// </summary>
            public const string Num_ThouDec2 = "#,##0.00_ ";

            /// <summary>
            /// 0->0.0
            /// </summary>
            public const string Num_Dec = "0.0_ ";

            /// <summary>
            /// 0->0.00
            /// </summary>
            public const string Num_Dec2 = "0.00_ ";
            /// <summary>
            /// 0->0.000
            /// </summary>
            public const string Num_Dec3 = "0.000_ ";
            /// <summary>
            /// 0->0.0000
            /// </summary>
            public const string Num_Dec4 = "0.0000_ ";

            /// <summary>
            /// Num Normal
            /// </summary>
            public const string Num_Normal = "0_ ";

            /// <summary>
            /// Percent
            /// </summary>
            public const string Percent_Dec2 = "0.00%";

            /// <summary>
            /// 百分比(1位小数)
            /// </summary>
            public const string Percent_Dec1 = "0.0%";

            /// <summary>
            /// 百分比(0位小数)
            /// </summary>
            public const string Percent_Dec0 = "0%";

            /// <summary>
            /// 字符串
            /// </summary>
            public const string CharString = "@";

            /// <summary>
            /// 通用
            /// </summary>
            public const string Common = "G/通用格式";

            /// <summary>
            /// 年月日期格式
            /// </summary>
            public const string ShortDateString = "yyyy" + "年" + "m" + "月" + ";@";

            /// <summary>
            /// 年月日日期格式
            /// </summary>
            public const string DefaultShortDate = "yyyy-mm-dd;@";

            /// <summary>
            /// 年月日日期格式
            /// </summary>
            public const string ShortDate = "yyyy/m/d";

            /// <summary>
            /// 年度日期格式
            /// </summary>
            public const string YearDateString = "yyyy" + "年" + ";@";
        }

        #endregion

        #region private method

        private static List<ExportRowItem> GetExportData<T>(this List<T> dataList, int dataRowIndex) where T : IExportModel
        {
            if (dataList.Count() == 0)
            {
                return new List<ExportRowItem>();
            }

            var members = typeof(T).GetProperties().Where(s => s.GetCustomAttribute(typeof(ColumnPropertyAttribute)) != null);

            var exportData = new List<ExportRowItem>();
            int firstRow = dataRowIndex;
            foreach (T dataItem in dataList)
            {
                var row = new ExportRowItem
                {
                    RowNo = firstRow,
                    Cols = new List<ExportColItem>()
                };

                foreach (PropertyInfo col in members)
                {
                    var attribute = (ColumnPropertyAttribute)col.GetCustomAttribute(typeof(ColumnPropertyAttribute));
                    if (attribute != null)
                    {
                        row.Cols.Add(new ExportColItem
                        {
                            ColNo = attribute.ColIndex,
                            Value = col.GetValue(dataItem)
                        });
                    }
                }

                exportData.Add(row);
                firstRow++;
            }

            return exportData;
        }

        internal class ExportRowItem
        {
            /// <summary>
            /// 
            /// </summary>
            public int RowNo { get; set; }

            /// <summary>
            /// 
            /// </summary>
            public List<ExportColItem> Cols { get; set; }
        }

        internal class ExportColItem
        {
            /// <summary>
            /// 
            /// </summary>
            public int ColNo { get; set; }

            /// <summary>
            /// 
            /// </summary>
            public object Value { get; set; }
        }

        #endregion
    }
}
