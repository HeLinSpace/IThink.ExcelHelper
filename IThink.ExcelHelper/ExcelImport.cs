using System;
using System.Collections.Generic;
using System.IO;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace H.Npoi.ExcelHelper
{
    /// <summary>
    /// Excel Helper for NPOI.
    /// </summary>
    public class ExcelImport : ImportSheetInfo, IDisposable
    {
        /// <summary>
        /// open workbook by workbook
        /// </summary>
        /// <param name="workbook"></param>
        public ExcelImport(IWorkbook workbook)
        {
            Workbook = workbook;
        }

        private List<SheetDataModel> _allSheetData { get; set; }

        /// <summary>
        /// all sheet original data.
        /// </summary>
        public List<SheetDataModel> AllSheetData
        {
            get
            {
                if (_allSheetData == null)
                {
                    ReadAllSheets();
                }

                return _allSheetData;
            }
        }

        /// <summary>
        /// open workbook by file
        /// </summary>
        /// <param name="path"></param>
        public ExcelImport(string path)
        {
            if (path.ToLower().EndsWith(".xlsx"))
            {
                Workbook = new XSSFWorkbook(path);
            }
            else
            {
                using (var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    Workbook = new HSSFWorkbook(fs);
                }
            }
        }

        /// <summary>
        /// open workbook by stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="isXlsx"></param>
        /// <returns></returns>
        public ExcelImport(Stream stream, bool isXlsx = true)
        {
            if (isXlsx)
            {
                Workbook = new XSSFWorkbook(stream);
            }
            else
            {
                Workbook = new HSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetNo">the index of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <returns></returns>
        public dynamic Import<T>(int sheetNo, int startIndex, int errorIndex, Func<List<T>, ImportSheetInfo, dynamic> business) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData<T>(sheetNo, startIndex, errorIndex);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetName">the name of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <returns></returns>
        public dynamic Import<T>(string sheetName, int startIndex, int errorIndex, Func<List<T>, ImportSheetInfo, dynamic> business) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData<T>(sheetName, startIndex, errorIndex);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheet">the sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <returns></returns>
        public dynamic Import<T>(ISheet sheet, int startIndex, int errorIndex, Func<List<T>, ImportSheetInfo, dynamic> business) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData<T>(sheet, startIndex, errorIndex);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetNo"></param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <returns></returns>
        public dynamic Import<T>(int sheetNo, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, Func<List<T>,ImportSheetInfo, dynamic> business) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData(sheetNo, startIndex, errorIndex, rowFunc);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <returns></returns>
        public dynamic Import<T>(string sheetName, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, Func<List<T>, ImportSheetInfo, dynamic> business) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData(sheetName, startIndex, errorIndex, rowFunc);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheet"></param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <returns></returns>
        public dynamic Import<T>(ISheet sheet, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, Func<List<T>, ImportSheetInfo, dynamic> business) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData(sheet, startIndex, errorIndex, rowFunc);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetNo">the index of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <param name="extraParam">提供给 rowFunc or business 的额外参数</param>
        /// <returns></returns>
        public dynamic Import<T>(int sheetNo, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, Func<List<T>, ImportSheetInfo, dynamic> business, dynamic extraParam) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData(sheetNo, startIndex, errorIndex, rowFunc, extraParam);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetName">the name of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <param name="extraParam">提供给 rowFunc or business 的额外参数</param>
        /// <returns></returns>
        public dynamic Import<T>(string sheetName, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, Func<List<T>, ImportSheetInfo, dynamic> business, dynamic extraParam) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData(sheetName, startIndex, errorIndex, rowFunc, extraParam);

            return business(dataList, this);
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheet">the index of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="business">执行业务处理（执行于获取所有数据以后）</param>
        /// <param name="extraParam">提供给 rowFunc or business 的额外参数</param>
        /// <returns></returns>
        public dynamic Import<T>(ISheet sheet, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, Func<List<T>, ImportSheetInfo, dynamic> business, dynamic extraParam) where T : ImportBaseModel, new()
        {
            var dataList = GetSheetData(sheet, startIndex, errorIndex, rowFunc, extraParam);

            return business(dataList, this);
        }

        private void ReadAllSheets() 
        {
            var sheetCount = Workbook.NumberOfSheets;

            for (var index = 0; index < sheetCount; index++) 
            {
                var sheetData = GetSheetData(index);
                _allSheetData.Add(sheetData);
            }
        }
    }
}
