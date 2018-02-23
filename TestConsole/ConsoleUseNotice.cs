using ExcelNoneMerge;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    class ConsoleUseNotice : INoneMergeNotice
    {
        public void CloseNotice(string workBookPath)
        {
            Console.WriteLine(workBookPath + "关闭");
        }

        public void OpenNotice(string workBookPath)
        {
            Console.WriteLine(workBookPath + "打开");
        }

        public void SheetEndNotice(string sheetName)
        {
            Console.WriteLine(sheetName + "完成解组");
        }

        public void SheetStartNotice(string sheetName)
        {
            Console.WriteLine(sheetName + "开始解组");
        }
    }
}
