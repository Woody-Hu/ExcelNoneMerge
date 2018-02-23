
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelNoneMerge
{
    /// <summary>
    /// ExcelNoneMerge管理器
    /// </summary>
    internal class ExcelNoneMergeManger
    {
        /// <summary>
        /// 使用的Excel程序
        /// </summary>
        private Application m_thisApplication = null;

        /// <summary>
        /// 使用的通知器
        /// </summary>
        private INoneMergeNotice m_useNotice = null;

        /// <summary>
        /// 使用的最大Column值
        /// </summary>
        private uint m_useColumnMaxNumber = 20;

        /// <summary>
        /// 构造管理器
        /// </summary>
        /// <param name="inputNotice">使用的通知器</param>
        /// <param name="inputMaxColumnNumber">使用的最大Column号</param>
        public ExcelNoneMergeManger(INoneMergeNotice inputNotice = null,uint? inputMaxColumnNumber = null)
        {

            m_useNotice = inputNotice;
            m_thisApplication = new Application();
            m_thisApplication.Visible = false;
            m_thisApplication.DisplayAlerts = false;
            m_thisApplication.AlertBeforeOverwriting = false;

            if (inputMaxColumnNumber.HasValue)
            {
                m_useColumnMaxNumber = inputMaxColumnNumber.Value;
            }
        }

        /// <summary>
        /// 析构函数时关闭
        /// </summary>
        ~ExcelNoneMergeManger()
        {
            QuitApplication();
        }

        /// <summary>
        /// 关闭APP
        /// </summary>
        public void QuitApplication()
        {
            if (null != m_thisApplication)
            {
                m_thisApplication.Quit();
                m_thisApplication = null;
            }
        }

        /// <summary>
        /// 将输入工作簿解除合并单元格
        /// </summary>
        /// <param name="inputPath"></param>
        public void WorkBookNoneMergeCell(string inputPath)
        {
            var tempWorkBook = m_thisApplication.Workbooks.Open(inputPath);
            if (null != m_useNotice)
            {
                m_useNotice.OpenNotice(inputPath);
            }

            foreach (var oneSheet in tempWorkBook.Worksheets)
            {
                Worksheet tempSheet = oneSheet as Worksheet;
                if (null != m_useNotice)
                {
                    m_useNotice.SheetStartNotice(tempSheet.Name);
                }
                NoneMergeCell(tempSheet);
                if (null != m_useNotice)
                {
                    m_useNotice.SheetEndNotice(tempSheet.Name);
                }
            }
            
            tempWorkBook.Save();
            tempWorkBook.Close();

            if (null != m_useNotice)
            {
                m_useNotice.CloseNotice(inputPath);
            }
            return;
        }

        /// <summary>
        /// 解除合并单元格一个工作表
        /// </summary>
        /// <param name="inputSheet"></param>
        private void NoneMergeCell(Worksheet inputSheet)
        {
            List<Range> lstRangeNeedNoneMerge = new List<Range>();

            Dictionary<Range, List<Range>> dicMergeRangeGroup = new Dictionary<Range, List<Range>>(new RangeComparer());

            int useEnd = 0;

            useEnd = FindMaxRow(inputSheet, useEnd);

            FindMergedCell(inputSheet, lstRangeNeedNoneMerge, useEnd);

            GroupCell(lstRangeNeedNoneMerge, dicMergeRangeGroup);

            UnMergeCell(inputSheet, dicMergeRangeGroup);

        }

        /// <summary>
        /// 解除一组单元格合并
        /// </summary>
        /// <param name="inputSheet">输入的Excel表</param>
        /// <param name="dicMergeRangeGroup">成组的Cell</param>
        private void UnMergeCell(Worksheet inputSheet, Dictionary<Range, List<Range>> dicMergeRangeGroup)
        {
            foreach (var oneMergeKVP in dicMergeRangeGroup)
            {
                int useRow = oneMergeKVP.Key.Row;
                int useColumn = oneMergeKVP.Key.Column;
                Range useRange = inputSheet.Cells[useRow, useColumn];
                string mergeValue = useRange.Text;
                oneMergeKVP.Key.UnMerge();

                foreach (var oneRange in oneMergeKVP.Value)
                {
                    oneRange.Value = mergeValue;
                }
            }
        }

        /// <summary>
        /// 将合并单元格的Cell解组
        /// </summary>
        /// <param name="lstRangeNeedNoneMerge"></param>
        /// <param name="dicMergeRangeGroup"></param>
        private void GroupCell(List<Range> lstRangeNeedNoneMerge, Dictionary<Range, List<Range>> dicMergeRangeGroup)
        {
            foreach (var oneRange in lstRangeNeedNoneMerge)
            {
                if (!dicMergeRangeGroup.ContainsKey(oneRange.MergeArea))
                {
                    dicMergeRangeGroup.Add(oneRange.MergeArea, new List<Range>());
                }
                dicMergeRangeGroup[oneRange.MergeArea].Add(oneRange);
            }
        }

        /// <summary>
        /// 寻找被合并单元格的Cell
        /// </summary>
        /// <param name="inputSheet"></param>
        /// <param name="lstRangeNeedNoneMerge"></param>
        /// <param name="useEnd"></param>
        private void FindMergedCell(Worksheet inputSheet, List<Range> lstRangeNeedNoneMerge, int useEnd)
        {
            for (int oneColumnIndex = 1; oneColumnIndex < m_useColumnMaxNumber + 1; oneColumnIndex++)
            {
                for (int oneRowIndex = 1; oneRowIndex < useEnd; oneRowIndex++)
                {
                    Range tempCell = inputSheet.Cells[oneRowIndex, oneColumnIndex];

                    string oneString = tempCell.Text;

                    if (true == tempCell.MergeCells)
                    {
                        lstRangeNeedNoneMerge.Add(tempCell);
                    }
                }
            }
        }

        /// <summary>
        /// 寻找最大行数
        /// </summary>
        /// <param name="inputSheet"></param>
        /// <param name="useEnd"></param>
        /// <returns></returns>
        private int FindMaxRow(Worksheet inputSheet, int useEnd)
        {
            for (int oneColumnIndex = 1; oneColumnIndex < m_useColumnMaxNumber + 1; oneColumnIndex++)
            {
                Range tempColumn = inputSheet.Columns[oneColumnIndex];
                int endRow = tempColumn.Rows[65536].End[XlDirection.xlUp].Row + 1;

                useEnd = Math.Max(useEnd, endRow);

            }

            return useEnd;
        }
    }
}
