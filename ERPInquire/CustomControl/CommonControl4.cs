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

namespace ERPInquire.CustomControl
{
    public partial class CommonControl4 : UserControl
    {
        #region 定量

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
        DataTable dt = new DataTable();
        DateTime starttime;
        DateTime endtime;

        #endregion

        #region 初始化
        public CommonControl4()
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
            starttime = this.sdate.Value;
            endtime = this.edate.Value;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(4, sdate, toolStrip1);
            MyMIS.AddToolstrip.AddDTPtoToolstrip(6, edate, toolStrip1);
        }

        private void CommonControl4_Load(object sender, EventArgs e)
        {
            toolStripLabel2.Text = TimeKey;
        }
        #endregion

        #region 窗体事件

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < e.RowCount; i++)
                this.dataGridView1.Rows[e.RowIndex + i].HeaderCell.Value = (e.RowIndex + i + 1).ToString();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
         //   MyMIS.ExcleIO.TableToExcelForXLSX2003(NodeName, dt);
        }
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
           MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            if (toolStripTextBox1.Text.Equals(""))
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
                                     new SqlParameter("@STARTTIME",SqlDbType.DateTime),
                                     new SqlParameter("@ENDTIME",SqlDbType.DateTime)
              };
                        para[0].Value = sdate.Value.ToString("yyyy-MM-dd") + " 00:00:00";
                        para[1].Value = edate.Value.ToString("yyyy-MM-dd") + " 23:59:59";
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
                            toolStripLabel3.Text = "总行数:" + (dt.Rows.Count).ToString();
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
                toolStripLabel3.Text = "总行数:" + (dtt.Rows.Count).ToString();
            }
             catch
            {
                    MessageBox.Show("此列暂不支持筛选", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
        }
            this.Cursor = Cursors.Default;//正常状态
        }
        #endregion

        #region 方法

        #endregion

        #region 分析
        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }
        #endregion

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
