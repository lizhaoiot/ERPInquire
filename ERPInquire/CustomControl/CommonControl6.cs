using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Com.Hui.iMRP.Utils;
using SqlConnect;
using Com.iCost;

namespace ERPInquire.CustomControl
{
    public partial class CommonControl6 : UserControl
    {
        #region 定量

        public DataTable m_table = null;
        public string ProName;
        public string NodeName;
        public string TimeKey;
        #endregion

        #region 变量
        DateTimePicker sdate;
        DateTimePicker edate;
        TextBox filter;
        public string sqlstr;
        public ConnectStr connect;
        DateTime starttime;
        DateTime endtime;
        DataTable dt = new DataTable();
        #endregion

        #region 初始化
        public CommonControl6()
        {
            InitializeComponent();
            Initialize();
            this.dataGridView1.RowsAdded += new DataGridViewRowsAddedEventHandler(dataGridView1_RowsAdded);
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行被选中
        }
        private void Initialize()
        {
            DateTime dt = DateTime.Today;
            this.sdate = new DateTimePicker();
            this.edate = new DateTimePicker();
            this.sdate.Value = dt.AddDays(-30d);
            this.edate.Value = dt;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(4, sdate, toolStrip1);
            MyMIS.AddToolstrip.AddDTPtoToolstrip(6, edate, toolStrip1);
        }
        private void CommonControl6_Load(object sender, EventArgs e)
        {
            toolStripComboBox1.SelectedIndex = 0;
            toolStripLabel2.Text = TimeKey;
        }
        #endregion

        #region 窗体事件
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < e.RowCount; i++)
                this.dataGridView1.Rows[e.RowIndex + i].HeaderCell.Value = (e.RowIndex + i + 1).ToString();
        }
        //查询
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            if (toolStripTextBox1.Text.Equals(""))
            {
                starttime = Convert.ToDateTime(sdate.Value.ToString("yyyy-MM-dd") + " 0:00:00");
                endtime = Convert.ToDateTime(edate.Value.ToString("yyyy-MM-dd") + " 23:59:59");
                if (toolStripComboBox1.SelectedIndex == 0)
                {
                    dt = DataProvider.GetYsDataTable(starttime.Date, endtime);
                    this.dataGridView1.DataSource = dt;
                       
                }
                else
                {
                     dt= DataProvider.GetYhCpDataTable(starttime, endtime);
                    this.dataGridView1.DataSource = dt;
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
                int i = toolStripComboBox2.ComboBox.SelectedIndex;
                string dtcolounname = toolStripComboBox2.SelectedItem.ToString();
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
        //导出EXCEL
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
            helper.DataTableToExcel(this.dataGridView1.DataSource as DataTable,"sheet1", true);
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
