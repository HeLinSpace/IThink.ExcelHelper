using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NPOI.SS.UserModel;

namespace H.Npoi.ExcelHelper
{
    /// <summary>
    /// Excel Helper for NPOI.
    /// </summary>
    public class ImportSheetInfo
    {
        private ICellStyle _defaultCellStyle;

        /// <summary>
        /// current sheet original data.
        /// </summary>
        public List<ImportDataModel> SheetData { get; private set; }

        /// <summary>
        /// current workbook
        /// </summary>
        public IWorkbook Workbook { get; internal set; }

        /// <summary>
        /// current worksheet
        /// </summary>
        public ISheet CurrentSheet { get; internal set; }

        /// <summary>
        /// the data row start number
        /// </summary>
        public int DataStartIndex { get; internal set; }

        /// <summary>
        /// error message col
        /// </summary>
        public int ErrorColIndex { get; internal set; }

        /// <summary>
        /// last row contained n this sheet (0-based)
        /// </summary>
        public int LastRowNum { get; internal set; }

        /// <summary>
        /// params for row func or business func
        /// </summary>
        public dynamic ExtraParam { get; set; }

        /// <summary>
        /// current worksheet
        /// </summary>
        public ISheet GetSheetAt(int sheetNo)
        {
            return Workbook.GetSheetAt(sheetNo);
        }

        /// <summary>
        /// current worksheet
        /// </summary>
        public ISheet GetSheet(string sheetName)
        {
            return Workbook.GetSheet(sheetName);
        }

        /// <summary>
        /// get or set default cellStyle for current workbook before your own operate.
        /// </summary>
        public ICellStyle DefaultCellStyle
        {
            get
            {
                if (_defaultCellStyle == null)
                {
                    return NExcelHelper.CreateCellStyle(Workbook);
                }

                return _defaultCellStyle;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("DefaultCellStyle");
                }
                _defaultCellStyle = value;
            }
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetNo">the index of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <returns></returns>
        public List<T> GetSheetData<T>(int sheetNo, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc = null) where T : ImportBaseModel, new()
        {
            if (Workbook == null)
            {
                throw new Exception("The workbook is not initialized,please set the value or call method IExcelHelper.Open before.");
            }

            CurrentSheet = Workbook.GetSheetAt(sheetNo);

            return GetSheetData(startIndex, errorIndex, rowFunc);
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetName">the name of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <returns></returns>
        public List<T> GetSheetData<T>(string sheetName, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc = null) where T : ImportBaseModel, new()
        {
            if (Workbook == null)
            {
                throw new Exception("The workbook is not initialized,please set the value or call method IExcelHelper.Open before.");
            }

            CurrentSheet = Workbook.GetSheet(sheetName);

            return GetSheetData(startIndex, errorIndex, rowFunc);
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="worksheet">current sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <returns></returns>
        public List<T> GetSheetData<T>(ISheet worksheet, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc = null) where T : ImportBaseModel, new()
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException("Worksheet");
            }

            CurrentSheet = worksheet;

            return GetSheetData(startIndex, errorIndex, rowFunc);
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetNo">the index of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="extraParam">额外参数</param>
        /// <returns></returns>
        public List<T> GetSheetData<T>(int sheetNo, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, dynamic extraParam) where T : ImportBaseModel, new()
        {
            if (Workbook == null)
            {
                throw new Exception("The workbook is not initialized,please set the value or call method IExcelHelper.Open before.");
            }

            CurrentSheet = Workbook.GetSheetAt(sheetNo);

            return GetSheetData(startIndex, errorIndex, rowFunc, extraParam);
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="sheetName">the name of sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="extraParam">额外参数</param>
        /// <returns></returns>
        public List<T> GetSheetData<T>(string sheetName, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, dynamic extraParam) where T : ImportBaseModel, new()
        {
            if (Workbook == null)
            {
                throw new Exception("The workbook is not initialized,please set the value or call method IExcelHelper.Open before.");
            }

            CurrentSheet = Workbook.GetSheet(sheetName);

            return GetSheetData(startIndex, errorIndex, rowFunc, extraParam);
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="worksheet">current sheet</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="extraParam">提供给 rowFunc 的额外参数</param>
        /// <returns></returns>
        public List<T> GetSheetData<T>(ISheet worksheet, int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, dynamic extraParam) where T : ImportBaseModel, new()
        {
            CurrentSheet = worksheet;

            return GetSheetData(startIndex, errorIndex, rowFunc, extraParam);
        }

        /// <summary>
        /// 写入错误文件
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="errRowList">写入数据源</param>
        /// <param name="errorColName">错误列列名</param>
        /// <param name="cellStyle">错误列格式</param>
        /// <returns></returns>
        public void WriteError<T>(List<T> errRowList, string errorColName = "Error", ICellStyle cellStyle = null) where T : ImportBaseModel
        {
            if (cellStyle == null)
            {
                cellStyle = _defaultCellStyle;
            }

            //写错误列标题
            var errorHead = CurrentSheet.GetRow(DataStartIndex - 1).CreateCell(ErrorColIndex);
            errorHead.SetCellValue(errorColName);
            errorHead.CellStyle = cellStyle;

            //保持和导入的文件中的排序一致
            errRowList = errRowList.OrderBy(o => o.RowNo).ToList();

            // 写错误原因数据
            foreach (var item in errRowList)
            {
                var errorCell = CurrentSheet.GetRow(item.RowNo).CreateCell(ErrorColIndex);
                errorCell.SetCellValue(item.ErrorMsg);
                errorCell.CellStyle = cellStyle;
            }
        }

        /// <summary>
        /// 写入错误文件
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="errRowList">写入数据源</param>
        /// <param name="filePath">保存文件名</param>
        /// <param name="errorColName">错误列列名</param>
        /// <param name="cellStyle">错误列格式</param>
        /// <returns></returns>
        public void WriteErrorFile<T>(List<T> errRowList, string filePath, string errorColName = "Error", ICellStyle cellStyle = null) where T : ImportBaseModel
        {
            if (cellStyle == null)
            {
                cellStyle = _defaultCellStyle;
            }

            //写错误列标题
            var errorHead = CurrentSheet.GetRow(DataStartIndex - 1).CreateCell(ErrorColIndex);
            errorHead.SetCellValue(errorColName);
            errorHead.CellStyle = cellStyle;

            //保持和导入的文件中的排序一致
            errRowList = errRowList.OrderBy(o => o.RowNo).ToList();

            // 写错误原因数据
            foreach (var item in errRowList)
            {
                var errorCell = CurrentSheet.GetRow(item.RowNo).CreateCell(ErrorColIndex);
                errorCell.SetCellValue(item.ErrorMsg);
                errorCell.CellStyle = cellStyle;
            }

            Workbook.Save(filePath);
        }

        /// <summary>
        /// 写入错误文件
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="worksheet">工作表</param>
        /// <param name="errRowList">写入数据源</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列</param>
        /// <param name="errorColName">错误列列名</param>
        /// <param name="cellStyle">错误列格式</param>
        /// <returns></returns>
        public ISheet WriteError<T>(ISheet worksheet, List<T> errRowList, int startIndex, int errorIndex, string errorColName = "Error", ICellStyle cellStyle = null) where T : ImportBaseModel
        {
            if (cellStyle == null)
            {
                cellStyle = _defaultCellStyle;
            }

            //写错误列标题
            var errorHead = worksheet.GetRow(startIndex - 1).CreateCell(errorIndex);
            errorHead.SetCellValue(errorColName);
            errorHead.CellStyle = cellStyle;

            //保持和导入的文件中的排序一致
            errRowList = errRowList.OrderBy(o => o.RowNo).ToList();

            // 写错误原因数据
            foreach (var item in errRowList)
            {
                var errorCell = worksheet.GetRow(item.RowNo).CreateCell(errorIndex);
                errorCell.SetCellValue(item.ErrorMsg);
                errorCell.CellStyle = cellStyle;
            }

            return worksheet;
        }

        /// <summary>
        /// 写入错误文件
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="worksheet">工作表</param>
        /// <param name="errRowList">写入数据源</param>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列</param>
        /// <param name="filePath">保存文件名</param>
        /// <param name="errorColName">错误列列名</param>
        /// <param name="cellStyle">错误列格式</param>
        /// <returns></returns>
        public void WriteErrorFile<T>(ISheet worksheet, List<T> errRowList, int startIndex, int errorIndex, string filePath, string errorColName = "Error", ICellStyle cellStyle = null) where T : ImportBaseModel
        {
            if (cellStyle == null)
            {
                cellStyle = _defaultCellStyle;
            }

            //写错误列标题
            var errorHead = worksheet.GetRow(startIndex - 1).CreateCell(errorIndex);
            errorHead.SetCellValue(errorColName);
            errorHead.CellStyle = cellStyle;

            //保持和导入的文件中的排序一致
            errRowList = errRowList.OrderBy(o => o.RowNo).ToList();

            // 写错误原因数据
            foreach (var item in errRowList)
            {
                var errorCell = worksheet.GetRow(item.RowNo).CreateCell(errorIndex);
                errorCell.SetCellValue(item.ErrorMsg);
                errorCell.CellStyle = cellStyle;
            }

            worksheet.Workbook.Save(filePath);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            if (Workbook != null)
            {
                Workbook.Close();
            }

            _defaultCellStyle = null;
            SheetData = null;
            GC.Collect();
            GC.SuppressFinalize(this);
        }

        #region private method

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T">接收数据的模型，需从Excel中获取的属性必须标记<see cref="ColumnPropertyAttribute"/></typeparam>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <returns></returns>
        private List<T> GetSheetData<T>(int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc = null) where T : ImportBaseModel, new()
        {
            if (CurrentSheet == null)
            {
                throw new ArgumentNullException("Worksheet");
            }

            // 获得总行数
            int rowTotalCount = CurrentSheet.LastRowNum + 1;

            if (rowTotalCount < startIndex)
            {
                throw new InvalidOperationException("start index is greater than the sheet's total count.");
            }

            DataStartIndex = startIndex;
            ErrorColIndex = errorIndex;
            LastRowNum = CurrentSheet.LastRowNum;

            var dataList = new List<T>();
            SheetData = new List<ImportDataModel>();

            var members = typeof(T).GetProperties()
                .Select(s => (ColumnPropertyAttribute)s.GetCustomAttribute(typeof(ColumnPropertyAttribute)))
                .Where(s => s != null)
                .Select(s => s.ColIndex).OrderBy(s => s).ToList();

            // 读取数据
            for (var rowIdx = startIndex; rowIdx < rowTotalCount; rowIdx++)
            {
                // 行数据
                var row = CurrentSheet.GetRow(rowIdx);
                var rowData = new ImportDataModel
                {
                    RowNo = rowIdx,
                    Row = new List<ImportColumnModel>()
                };

                if (row != null)
                {
                    foreach (var colIdx in members)
                    {
                        var value = row.GetCell(colIdx)?.GetValue();
                        rowData.Row.Add(new ImportColumnModel { ColIndex = colIdx, Value = value });
                    }
                }

                // 空行直接跳过
                if (!rowData.Row.Any(s => s.Value != null))
                {
                    continue;
                }

                SheetData.Add(rowData);
                var dataItem = rowData.ToModel<T>();

                if (rowFunc != null)
                {
                    // 格式检查
                    rowFunc(dataItem, this);
                }

                dataList.Add(dataItem);
            }

            return dataList;
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="startIndex">数据行</param>
        /// <param name="errorIndex">错误列（最后一列数据+1）</param>
        /// <param name="rowFunc">行数据获取后执行，常用于数据检查</param>
        /// <param name="extraParam">额外参数</param>
        /// <returns></returns>
        private List<T> GetSheetData<T>(int startIndex, int errorIndex, Action<T, ImportSheetInfo> rowFunc, dynamic extraParam) where T : ImportBaseModel, new()
        {
            if (CurrentSheet == null)
            {
                throw new ArgumentNullException("Worksheet");
            }

            // 获得总行数
            int rowTotalCount = CurrentSheet.LastRowNum + 1;

            if (rowTotalCount < startIndex)
            {
                throw new InvalidOperationException("start index is greater than the sheet's total count.");
            }

            DataStartIndex = startIndex;
            ErrorColIndex = errorIndex;
            ExtraParam = extraParam;
            LastRowNum = CurrentSheet.LastRowNum;

            var dataList = new List<T>();
            SheetData = new List<ImportDataModel>();

            var members = typeof(T).GetProperties()
                .Select(s => (ColumnPropertyAttribute)s.GetCustomAttribute(typeof(ColumnPropertyAttribute)))
                .Where(s => s != null)
                .Select(s => s.ColIndex).OrderBy(s => s).ToList();

            // 读取数据
            for (var rowIdx = startIndex; rowIdx < rowTotalCount; rowIdx++)
            {
                // 行数据
                var row = CurrentSheet.GetRow(rowIdx);
                var rowData = new ImportDataModel
                {
                    RowNo = rowIdx,
                    Row = new List<ImportColumnModel>()
                };

                if (row != null) 
                {
                    foreach (var colIdx in members)
                    {
                        var value = row.GetCell(colIdx)?.GetValue();
                        rowData.Row.Add(new ImportColumnModel { ColIndex = colIdx, Value = value });
                    }
                }

                // 空行直接跳过
                if (!rowData.Row.Any(s => s.Value != null))
                {
                    continue;
                }

                SheetData.Add(rowData);
                var dataItem = rowData.ToModel<T>();

                if (rowFunc != null)
                {
                    // 格式检查
                    rowFunc(dataItem, this);
                }

                dataList.Add(dataItem);
            }

            return dataList;
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <param name="sheetNo"></param>
        /// <returns></returns>
        internal SheetDataModel GetSheetData(int sheetNo)
        {
            CurrentSheet = Workbook.GetSheetAt(sheetNo);

            if (CurrentSheet == null)
            {
                throw new ArgumentNullException("Worksheet");
            }

            var sheetData = new SheetDataModel
            {
                SheetName = CurrentSheet.SheetName,
                SheetNo = sheetNo,
                Rows = new List<SheetDataRow>()
            };

            // 获得总行数
            int rowTotalCount = CurrentSheet.LastRowNum + 1;

            LastRowNum = CurrentSheet.LastRowNum;

            // 读取数据
            for (var rowIdx = 0; rowIdx < rowTotalCount; rowIdx++)
            {
                // 行数据
                var row = CurrentSheet.GetRow(rowIdx);

                var firstCellNum = row.FirstCellNum;
                var lastCellNum = row.LastCellNum;
                var rowData = new SheetDataRow
                {
                    RowIndex = rowIdx,
                    Columns = new List<SheetDataColumn>()
                };

                if (row != null)
                {
                    for (var colIdx = firstCellNum; colIdx <= lastCellNum; colIdx++)
                    {
                        var value = row.GetCell(colIdx)?.GetCellValue() ?? new SheetDataColumn { ColIndex = colIdx ,ValueType = ValueType.None };
                        rowData.Columns.Add(value);
                    }
                }

                // 空行直接跳过
                if (!rowData.Columns.Any(s => s.Value != null))
                {
                    continue;
                }

                sheetData.Rows.Add(rowData);
            }

            return sheetData;
        }

        #endregion
    }
}
