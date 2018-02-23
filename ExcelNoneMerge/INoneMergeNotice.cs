using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNoneMerge
{
    /// <summary>
    /// 通知器
    /// </summary>
    public interface INoneMergeNotice
    {
        /// <summary>
        /// 文件打开通知
        /// </summary>
        /// <param name="workBookPath"></param>
        void OpenNotice(string workBookPath);

        /// <summary>
        /// 文件关闭通知
        /// </summary>
        /// <param name="workBookPath"></param>
        void CloseNotice(string workBookPath);

        /// <summary>
        /// 表格打开通知
        /// </summary>
        /// <param name="sheetName"></param>
        void SheetStartNotice(string sheetName);

        /// <summary>
        /// 表格关闭通知
        /// </summary>
        /// <param name="sheetName"></param>
        void SheetEndNotice(string sheetName);
    }
}
