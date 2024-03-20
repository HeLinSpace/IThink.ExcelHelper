using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace H.Npoi.ExcelHelper
{
    /// <summary>
    /// Excel Helper for NPOI.
    /// </summary>
    public class ReadExcel : ReadSheetInfo, IDisposable
    {
        /// <summary>
        /// open workbook by workbook
        /// </summary>
        /// <param name="workbook"></param>
        public ReadExcel(IWorkbook workbook)
        {
            Workbook = workbook;
        }

        private List<SheetDataModel> _allSheetData { get; set; }

        /// <summary>
        /// all sheet original data.
        /// </summary>
        /// <param name="autoTransferDateValue">自动识别日期型数据 转换为Datetime</param>
        /// <returns></returns>
        public List<SheetDataModel> GetAllSheetData(bool autoTransferDateValue = false)
        {
            AutoTransferDateValue = autoTransferDateValue;

            if (_allSheetData == null)
            {
                ReadAllSheets();
            }

            return _allSheetData;
        }

        /// <summary>
        /// all sheet original data.
        /// </summary>
        /// <param name="sheetNo">the index of sheet</param>
        /// <param name="autoTransferDateValue">自动识别日期型数据 转换为Datetime</param>
        /// <returns></returns>
        public SheetDataModel GetSheetData(int sheetNo, bool autoTransferDateValue = false)
        {
            AutoTransferDateValue = autoTransferDateValue;

            if (_allSheetData != null)
            {
                return _allSheetData.FirstOrDefault(s => s.SheetNo == sheetNo);
            }

            return base.GetSheetData(sheetNo);
        }

        /// <summary>
        /// open workbook by file
        /// </summary>
        /// <param name="path"></param>
        public ReadExcel(string path)
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
        public ReadExcel(Stream stream, bool isXlsx = true)
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
        /// 
        /// </summary>
        public override void Dispose()
        {
            if (Workbook != null)
            {
                Workbook.Close();
            }

            CurrentSheet = null;
            CurrentSheetNo = null;
            CurrentSheetName = null;
            GC.Collect();
            GC.SuppressFinalize(this);
        }

        private void ReadAllSheets() 
        {
            _allSheetData = new List<SheetDataModel>();
            var sheetCount = Workbook.NumberOfSheets;

            for (var index = 0; index < sheetCount; index++) 
            {
                var sheetData = base.GetSheetData(index);
                _allSheetData.Add(sheetData);
            }
        }
    }
}
