using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNoneMerge
{
    /// <summary>
    /// 单元格相等比较器
    /// </summary>
    internal class RangeComparer : IEqualityComparer<Range>
    {
        /// <summary>
        /// 相等比较
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public bool Equals(Range x, Range y)
        {
            return x.Row == y.Row && x.Column == y.Column;
        }

        /// <summary>
        /// 哈希值方法
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int GetHashCode(Range obj)
        {
            string tempString = string.Empty;
            tempString = tempString + "row:" + obj.Row.ToString() + "column:" + obj.Column.ToString();
            return tempString.GetHashCode();
        }
    }
}
