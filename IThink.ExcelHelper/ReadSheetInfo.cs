using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace H.Npoi.ExcelHelper
{
    /// <summary>
    /// Excel Helper for NPOI.
    /// </summary>
    public class ReadSheetInfo
    {
        /// <summary>
        /// current workbook
        /// </summary>
        public IWorkbook Workbook { get; internal set; }

        /// <summary>
        /// current worksheet
        /// </summary>
        public ISheet CurrentSheet { get; internal set; }

        /// <summary>
        /// index of the sheet
        /// </summary>
        public int? CurrentSheetNo { get; internal set; }

        /// <summary>
        /// name of the sheet
        /// </summary>
        public string CurrentSheetName { get; internal set; }

        /// <summary>
        /// last row contained n this sheet (0-based)
        /// </summary>
        public int CurrentLastRowNum { get; internal set; }

        internal bool AutoTransferDateValue { get;  set; }

        /// <summary>
        /// current worksheet
        /// </summary>
        public ISheet GetSheetAt(int sheetNo)
        {
            CurrentSheet = Workbook.GetSheetAt(sheetNo);
            CurrentSheetNo = sheetNo;
            CurrentSheetName = CurrentSheet.SheetName;

            return CurrentSheet;
        }

        /// <summary>
        /// current worksheet
        /// </summary>
        public ISheet GetSheet(string sheetName)
        {
            CurrentSheetNo = Workbook.GetSheetIndex(sheetName);
            CurrentSheet = Workbook.GetSheet(sheetName);
            CurrentSheetName = sheetName;

            return CurrentSheet;
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <param name="sheetNo"></param>
        /// <returns></returns>
        public SheetDataModel GetSheetData(int sheetNo)
        { 
            CurrentSheet = Workbook.GetSheetAt(sheetNo);
            CurrentSheetNo = sheetNo;
            CurrentSheetName = CurrentSheet.SheetName;

            return GetSheetData();
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public SheetDataModel GetSheetData(string sheetName)
        {
            CurrentSheet = Workbook.GetSheet(sheetName);
            CurrentSheetNo = Workbook.GetSheetIndex(sheetName);
            CurrentSheetName = sheetName;

            return GetSheetData();
        }

        /// <summary>
        /// 
        /// </summary>
        public virtual void Dispose()
        {
            CurrentSheet = null;
            CurrentSheetNo = null;
            CurrentSheetName = null;
            GC.Collect();
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <returns></returns>
        private SheetDataModel GetSheetData()
        {
            if (CurrentSheet == null)
            {
                throw new ArgumentNullException("Worksheet");
            }

            var sheetData = new SheetDataModel
            {
                SheetName = CurrentSheetName,
                SheetNo = CurrentSheetNo.Value,
                Rows = new List<SheetDataRow>()
            };

            // 获得总行数
            int rowTotalCount = CurrentSheet.LastRowNum + 1;

            CurrentLastRowNum = CurrentSheet.LastRowNum;

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
                        var value = row.GetCell(colIdx)?.GetCellValue(AutoTransferDateValue) ?? new SheetDataColumn { ColIndex = colIdx, ValueType = ValueType.None };
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
    }
}
