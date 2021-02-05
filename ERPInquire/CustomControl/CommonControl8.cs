using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IO;
using SqlConnect;
using System.Data.SqlClient;
using ExcelOperation;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ERPInquire.CustomControl
{
    public partial class CommonControl8 : UserControl
    {
        #region 定量

        public string ProName;
        public string TheKeyword;
        public string NodeName;
        public string TimeKey;
        public string TheKeyword1;

        #endregion

        #region 变量
        DateTimePicker sdate;
        DateTimePicker edate;
        public string sqlstr;
        public ConnectStr connect;
        DataTable dt = new DataTable();
        #endregion

        #region 初始化
        public CommonControl8()
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
            this.sdate.Value = dt.AddDays(-2d);
            this.edate.Value = dt;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(4, sdate, toolStrip1);
            MyMIS.AddToolstrip.AddDTPtoToolstrip(6, edate, toolStrip1);
        }
        private void CommonControl8_Load(object sender, EventArgs e)
        {
            toolStripLabel2.Text = TheKeyword;
            toolStripLabel3.Text = TheKeyword1;
            toolStripLabel5.Text = TimeKey;
        }
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < e.RowCount; i++)
                this.dataGridView1.Rows[e.RowIndex + i].HeaderCell.Value = (e.RowIndex + i + 1).ToString();
        }

        #endregion

        #region 窗体事件
        //查询
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            if (toolStripTextBox3.Text.Equals(""))
            {
                dt.Clear();
                SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
                string s = connect.GetConnectStr("ERP");
                string connectionString = s;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = ProName;
                        SqlParameter[] para ={
                                     new SqlParameter("@DAH1",SqlDbType.VarChar),
                                     new SqlParameter("@DAH2",SqlDbType.VarChar),
                                     new SqlParameter("@STARTTIME",SqlDbType.DateTime),
                                     new SqlParameter("@ENDTIME",SqlDbType.DateTime)
              };
                        para[0].Value = toolStripTextBox1.Text;
                        para[1].Value = toolStripTextBox2.Text;
                        para[2].Value = sdate.Value.ToString("yyyy-MM-dd") + " 00:00:00";
                        para[3].Value = edate.Value.ToString("yyyy-MM-dd") + " 23:59:59";
                        try
                        {
                            cmd.CommandTimeout = 60 * 60 * 10000;
                            cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(dt);
                            conn.Close();
                            dataGridView1.DataSource = dt;
                            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
                            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
                            toolStripLabel6.Text = "总行数:" + (dt.Rows.Count).ToString();
                            dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
                            dataGridView1.Refresh();
                            //获得datatable的列名
                            List<string> ls = SqlHelper.GetColumnsByDataTable(dt);
                            toolStripComboBox1.ComboBox.DataSource = ls;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
            else
            {
                try {
                //获得现在所选列的列索引
                int i = toolStripComboBox1.ComboBox.SelectedIndex;
                string dtcolounname = toolStripComboBox1.SelectedItem.ToString();
                DataRow[] dr = dt.Select(dtcolounname + " like '%" + toolStripTextBox3.Text + "%'");
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
                toolStripLabel6.Text = "总行数:" + (dtt.Rows.Count).ToString();
            }
             catch
            {
                    MessageBox.Show("此列暂不支持筛选", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
        }
            this.Cursor = Cursors.Default;//正常状态
        }

        //导出EXECL
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }

        #endregion

        #region 方法

        #endregion

        private void toolStripTextBox3_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(toolStripTextBox3.Text))
            {
                toolStripTextBox2.Enabled = true;
                toolStripTextBox1.Enabled = true;
            }
            else
            {
                toolStripTextBox2.Enabled = false;
                toolStripTextBox1.Enabled = false;
            }
        }
    }
}
