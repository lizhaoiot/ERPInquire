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
namespace ERPInquire.CustomControl
{
    public partial class CommonControl13 : UserControl
    {
        #region 定量
         public string ProName;
         public string TheKeyword;
         public string NodeName;
        #endregion

        #region 变量
         DateTimePicker sdate;
         DateTimePicker edate;
         public string sqlstr;
         public ConnectStr connect;
         TextBox filter;
         DataTable dt = new DataTable();
        #endregion

        #region 初始化
         public CommonControl13()
        {
            InitializeComponent();
            Initialize();
            this.dataGridView1.RowsAdded += new DataGridViewRowsAddedEventHandler(dataGridView1_RowsAdded);
        }
         private void Initialize()
        {
            DateTime dt = DateTime.Today;
            this.sdate = new DateTimePicker();
            this.sdate.Value = dt;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(7, sdate, toolStrip1);
            this.edate = new DateTimePicker();
            this.edate.Value = dt;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(9, edate, toolStrip1);
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行被选中
        }
         private void CommonControl13_Load(object sender, EventArgs e)
        {
            toolStripComboBox1.SelectedIndex = 0;
            toolStripComboBox2.SelectedIndex = 0;
        }
        #endregion

        #region 窗体事件
         private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
         {
             for (int i = 0; i < e.RowCount; i++)
                 this.dataGridView1.Rows[e.RowIndex + i].HeaderCell.Value = (e.RowIndex + i + 1).ToString();
         }
        #endregion

        #region 方法
         //查询
         private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
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
                    cmd.CommandText = "AfterPrintingDaily3";
                    SqlParameter[] para ={
                                     new SqlParameter("@DAH1",SqlDbType.VarChar),
                                     new SqlParameter("@DAH2",SqlDbType.VarChar),
                                     new SqlParameter("@DAH3",SqlDbType.VarChar),
                                     new SqlParameter("@DAH4",SqlDbType.VarChar),
                                     new SqlParameter("@DAH5",SqlDbType.VarChar),
                                     new SqlParameter("@TIME1",SqlDbType.DateTime),
                                     new SqlParameter("@TIME2",SqlDbType.DateTime)
              };
                    para[0].Value = toolStripComboBox1.Text;//班组
                    para[1].Value = toolStripComboBox2.Text;//机台名称
                    para[2].Value = toolStripTextBox1.Text;//工单号
                    para[3].Value = toolStripTextBox2.Text;//产品名称
                    para[4].Value = toolStripTextBox3.Text;//部门名称
                    para[5].Value = sdate.Value.ToString("yyyy-MM-dd") + " 00:00:00";
                    para[6].Value = edate.Value.ToString("yyyy-MM-dd") + " 23:59:59";
                    try
                    {
                        cmd.CommandTimeout = 60 * 60 * 1000;
                        cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dt);
                        conn.Close();
                        dataGridView1.DataSource = dt;
                        dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
                        dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
                        dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
                        dataGridView1.Refresh();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            this.Cursor = Cursors.Default;//正常状态
        }
         //导出EXCEL
         private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }
         
        #endregion


    }
}
