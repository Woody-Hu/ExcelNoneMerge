using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNoneMerge
{
    /// <summary>
    /// 路径Excel解合并管理器
    /// </summary>
    public class DirectionExcelNoneMerge
    {
        /// <summary>
        /// 使用的通知器
        /// </summary>
        private INoneMergeNotice m_useNotice = null;

        /// <summary>
        /// 输入的路径
        /// </summary>
        private DirectoryInfo m_inputDirectory = null;

        /// <summary>
        /// 使用的Excel解组管理器
        /// </summary>
        private ExcelNoneMergeManger m_useManger = null;

        /// <summary>
        /// 所有的Excel文件全路径
        /// </summary>
        private List<string> m_lstPath = new List<string>();

        /// <summary>
        /// 特殊后缀
        /// </summary>
        private const string m_strXlsx = ".xlsx";

        /// <summary>
        /// 特殊后缀
        /// </summary>
        private const string m_strXls = ".xls";

        /// <summary>
        /// 构造管理器
        /// </summary>
        /// <param name="inputDirection">输入的路径</param>
        /// <param name="useNotice">使用的通知器</param>
        /// <param name="inputMaxColumnNumber">使用的最大Column号</param>
        public DirectionExcelNoneMerge(DirectoryInfo inputDirection, INoneMergeNotice useNotice = null, uint? inputMaxColumnNumber = null)
        {
            m_inputDirectory = inputDirection;
            m_useNotice = useNotice;
            m_useManger = new ExcelNoneMergeManger(m_useNotice, inputMaxColumnNumber);
            PreparePath(m_inputDirectory);
            m_lstPath = m_lstPath.Distinct().ToList();
        }

        /// <summary>
        /// 递归寻找所有Excel文件
        /// </summary>
        /// <param name="inputDirection"></param>
        private void PreparePath(DirectoryInfo inputDirection)
        {
            foreach (var oneFileInfo in inputDirection.GetFiles())
            {
                if (oneFileInfo.Extension.Equals(m_strXlsx) || oneFileInfo.Extension.Equals(m_strXls))
                {
                    m_lstPath.Add(oneFileInfo.FullName);
                }
            }

            foreach (var oneSubDirection in inputDirection.GetDirectories())
            {
                PreparePath(oneSubDirection);
            }
        }

        /// <summary>
        /// ExcelUnMerge
        /// </summary>
        public void UnMerge()
        {
            foreach (var onePath in m_lstPath)
            {
                m_useManger.WorkBookNoneMergeCell(onePath);
            }

            m_useManger.QuitApplication();
        }
    }
}
