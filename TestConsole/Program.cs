using ExcelNoneMerge;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputDir = Console.ReadLine();

            System.IO.DirectoryInfo useDir = new System.IO.DirectoryInfo(inputDir);

            DirectionExcelNoneMerge useManger = new DirectionExcelNoneMerge(useDir, new ConsoleUseNotice());
            useManger.UnMerge();

            Console.Read();
        }
    }
}
