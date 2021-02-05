using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SqlConnect;
using Com.Hui.iMRP.Utils;
using Com.iReport;

namespace ERPInquire.CustomControl
{
    public partial class CommonControl7 : UserControl
    {
        #region 定量

        public string NodeName;
        public string TimeKey;
        #endregion

        #region 变量
        DateTimePicker sdate; 
        DateTimePicker edate;
        TextBox filter;
        public string sqlstr;
        public ConnectStr connect;
        DataTable dt = new DataTable();
        DateTime starttime;
        DateTime endtime;
        #endregion

        #region 初始化
        public CommonControl7()
        {
            InitializeComponent();
            Initialize();
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行被选中
        }
        private void Initialize()
        {
            DateTime dt = DateTime.Today;
            this.sdate = new DateTimePicker();
            this.edate = new DateTimePicker();
            this.sdate.Value = dt.AddDays(-30d);
            this.edate.Value = dt;
            starttime = this.sdate.Value;
            endtime = this.edate.Value;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(4, sdate, toolStrip1);
            MyMIS.AddToolstrip.AddDTPtoToolstrip(6, edate, toolStrip1);
        }
        private void CommonControl7_Load(object sender, EventArgs e)
        {
            toolStripComboBox1.SelectedIndex = 0;
            toolStripLabel2.Text = TimeKey;
        }
        #endregion

        #region 窗体事件
        //查询
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            if (toolStripTextBox1.Text.Equals(""))
            {
                if (toolStripComboBox1.SelectedIndex == 0)
                {
                    this.dataGridView1.DataSource = DataProvider.GetDataTable(starttime, endtime, DataProvider.LB.单张书版票据);
                }
                else
                {
                    this.dataGridView1.DataSource = DataProvider.GetDataTable(starttime, endtime, DataProvider.LB.包装箱包装盒);
                }
                dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
                dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
                toolStripLabel4.Text = "总行数:" + (dataGridView1.RowCount).ToString();
                dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
                dataGridView1.Refresh();
                //获得datatable的列名
                List<string> ls = SqlHelper.GetColumnsByDataTable(dt);
                toolStripComboBox2.ComboBox.DataSource = ls;
            }
            else
            {
                try {
                //获得现在所选列的列索引
                int i = toolStripComboBox1.ComboBox.SelectedIndex;
                string dtcolounname = toolStripComboBox1.SelectedItem.ToString();
                DataRow[] dr = dt.Select(dtcolounname + " like '%" + toolStripTextBox1.Text + "%'");
                DataTable dtt = new DataTable();
                dtt = dt.Clone();
                if (dr.Length > 0)
                {
                    foreach (DataRow drVal in dr)
                    {
                        dtt.ImportRow(drVal);
                    }
                }
                dataGridView1.DataSource = dtt;
                toolStripLabel4.Text = "总行数:" + (dtt.Rows.Count).ToString();
            }
             catch
            {
                    MessageBox.Show("此列暂不支持筛选", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
        }
            this.Cursor = Cursors.Default;//等待
        }
        //导出execl
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }
        #endregion

        #region 方法
        private void Export2Excel()
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = NodeName + DateTime.Now.ToString("yyyyMMddhhmmss");
            dlg.Filter = "xlsx files(*.xlsx)|*.xlsx|xls files(*.xls)|*.xls|All files(*.*)|*.*";
            dlg.ShowDialog();
            if (dlg.FileName.IndexOf(":") < 0) return; //被点了"取消"
            ExcelHelper helper = new ExcelHelper(dlg.FileName);
            helper.DataTableToExcel(this.dataGridView1.DataSource as DataTable, "sheet1", true);
            MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(toolStripTextBox1.Text))
            {
                toolStripComboBox1.Enabled = true;
            }
            else
            {
                toolStripComboBox1.Enabled = false;
            }
        }
    }
}
