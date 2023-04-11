using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace H.Npoi.ExcelHelper
{
    /// <summary>
    /// 
    /// </summary>
    public class SheetDataModel
    {
        /// <summary>
        /// 
        /// </summary>
        public int SheetNo { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string SheetName { get; set; }

        /// <summary> 
        /// 行数据
        /// </summary>
        public List<SheetDataRow> Rows { get; set; }
    }

    /// <summary>
    /// 
    /// </summary>
    public class SheetDataRow
    {
        /// <summary> 
        /// 行号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary> 
        /// 列
        /// </summary>
        public List<SheetDataColumn> Columns { get; set; }
    }

    /// <summary>
    /// 
    /// </summary>
    public class SheetDataColumn
    {
        /// <summary> 
        /// 列号
        /// </summary>
        public int ColIndex { get; set; }

        /// <summary> 
        /// 值类型
        /// </summary>
        public ValueType ValueType { get; set; }

        /// <summary> 
        /// 单元格值
        /// </summary>
        public object Value { get; set; }
    }

    /// <summary>
    /// 
    /// </summary>
    public enum ValueType
    {
        /// <summary>
        /// 
        /// </summary>
        None,
        /// <summary>
        /// 
        /// </summary>
        String,
        /// <summary>
        /// 
        /// </summary>
        Boolean,
        /// <summary>
        /// 
        /// </summary>
        Numeric,
        /// <summary>
        /// 
        /// </summary>
        DateTime
    }
}
