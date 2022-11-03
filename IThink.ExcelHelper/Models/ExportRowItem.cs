using System;
using System.Collections.Generic;

namespace H.Npoi.ExcelHelper
{

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
}
