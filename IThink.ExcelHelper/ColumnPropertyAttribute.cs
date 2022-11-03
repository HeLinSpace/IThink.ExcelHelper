using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

namespace H.Npoi.ExcelHelper
{
    /// <summary>
    /// 
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnPropertyAttribute : Attribute
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="colIndex">列序号</param>
        public ColumnPropertyAttribute(int colIndex)
        {
            ColIndex = colIndex;
        }

        /// <summary>
        /// 
        /// </summary>
        public int ColIndex { get;private set; }
    }
}
