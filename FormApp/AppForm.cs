using ExcelNoneMerge;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FormApp
{
    public partial class AppForm : Form, INoneMergeNotice
    {
        public AppForm()
        {
            InitializeComponent();
        }

        public void CloseNotice(string workBookPath)
        {
            label_state.Text = workBookPath + "关闭";
        }

        public void OpenNotice(string workBookPath)
        {
            label_state.Text = workBookPath + "打开";
        }

        public void SheetEndNotice(string sheetName)
        {
            label_state.Text = sheetName + "完成";
        }

        public void SheetStartNotice(string sheetName)
        {
            label_state.Text = sheetName + "开始";
        }

        private void label_state_Click(object sender, EventArgs e)
        {

        }

        private void button_selectDirection_Click(object sender, EventArgs e)
        {
            string path = string.Empty;
            FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                path = fbd.SelectedPath;
            }

            button_selectDirection.Enabled = false;

            DirectionExcelNoneMerge useManger = new DirectionExcelNoneMerge(new System.IO.DirectoryInfo(path), this);

            useManger.UnMerge();

            MessageBox.Show("完成");

            button_selectDirection.Enabled = true;

            label_state.Text = "无";

        }
    }
}
