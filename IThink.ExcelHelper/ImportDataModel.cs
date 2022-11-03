using System.Collections.Generic;

namespace H.Npoi.ExcelHelper
{
    public class ImportDataModel : ImportBaseModel
    {
        /// <summary> 
        /// 行数据
        /// </summary>
        public List<ImportColumnModel> Row { get; set; }
    }

    public class ImportColumnModel
    {
        /// <summary> 
        /// 列号
        /// </summary>
        public int ColIndex { get; set; }

        /// <summary> 
        /// 单元格值
        /// </summary>
        public object Value { get; set; }
    }
}
