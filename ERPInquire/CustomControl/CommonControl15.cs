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
    public partial class CommonControl15 : UserControl
    {
        #region 定量
        DataTable dt = new DataTable();
        #endregion

        #region 变量
          public string NodeName;
          public string ProName;
        #endregion

        #region 初始化
        public CommonControl15()
        {
            InitializeComponent();
            this.dataGridView1.RowsAdded += new DataGridViewRowsAddedEventHandler(dataGridView1_RowsAdded);
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行被选中
        }
        #endregion

        #region 窗体事件
         private void CommonControl15_Load(object sender, EventArgs e)
        {
            #region 数据源初始化
              //物料大类
              /*DataTable dt1 = new DataTable();
              dt1 = SqlHelper.ExecuteDataTable("SELECT XH,WLLB FROM JCZL_WLLB_M");
              toolStripComboBox1.ComboBox.DataSource = dt1;
              toolStripComboBox1.ComboBox.ValueMember = "XH";
              toolStripComboBox1.ComboBox.DisplayMember = "WLLB";
              toolStripComboBox1.ComboBox.Text=""
              //物料子类
              DataTable dt2 = new DataTable();
              dt2 = SqlHelper.ExecuteDataTable("SELECT XH,WLZL FROM JCZL_WLLB_D");
              toolStripComboBox2.ComboBox.DataSource = dt2;
              toolStripComboBox2.ComboBox.ValueMember = "XH";
              toolStripComboBox2.ComboBox.DisplayMember = "WLZL";
              toolStripComboBox2.ComboBox.Text = "";*/
              toolStripComboBox4.SelectedIndex = 2;
            #endregion
        }
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
                    cmd.CommandText = ProName;
                    SqlParameter[] para ={
                    new SqlParameter("@WLMC",SqlDbType.VarChar),
                    new SqlParameter("@WLLB",SqlDbType.VarChar),
                    new SqlParameter("@ZBXH_CK",SqlDbType.VarChar),
                    new SqlParameter("@WLZL",SqlDbType.VarChar),
                    new SqlParameter("@WLFL",SqlDbType.VarChar),
                    new SqlParameter("@WLBH",SqlDbType.VarChar)
                 };
                    para[0].Value = toolStripTextBox2.Text;
                    para[1].Value = ConvertWLDL("");   
                    para[2].Value = ConvertCK(toolStripComboBox3.Text);
                    para[3].Value = ConvertWLZL("");
                    para[4].Value = ConvertWLFL(toolStripComboBox4.Text.ToString());
                    para[5].Value = toolStripTextBox1.Text;
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
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < e.RowCount; i++)
                this.dataGridView1.Rows[e.RowIndex + i].HeaderCell.Value = (e.RowIndex + i + 1).ToString();
        }
        #endregion

        #region 方法
        private string ConvertCK(string ck)
        {
          if (ck.Equals(""))
          {
              return "";
          }
          else
          {
              return SqlHelper.ExecuteDataTable("SELECT XH FROM JCZL_CKBM WHERE CKMC='"+ck+"'").Rows[0][0].ToString();
          }
        }
        private string ConvertWLDL(string WLDL)
        {
            if (WLDL.Equals(""))
            {
                return "";
            }
            else
            {
                return SqlHelper.ExecuteDataTable("SELECT XH FROM JCZL_WLLB_M WHERE WLLB='" + WLDL + "'").Rows[0][0].ToString();
            }
        }
        private string ConvertWLZL(string WLZL)
        {
            if (WLZL.Equals(""))
            {
                return "";
            }
            else
            {
                return SqlHelper.ExecuteDataTable("SELECT XH FROM JCZL_WLLB_D WHERE WLZL='" + WLZL + "'").Rows[0][0].ToString();
            }
        }
        private string ConvertWLFL(string WLZL)
        {
            if (WLZL.Equals(""))
            {
                return "1";
            }
            else if (WLZL.Equals("纸料"))
            {
                return "0";
            }
            else if (WLZL.Equals("辅料"))
            {
                return "1";
            }
            else
            {
                return "3";
            }
        }
        #endregion

    }
}
