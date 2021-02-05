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
    public partial class CommonControl14 : UserControl
    {
        #region 定量
         DataTable dt = new DataTable();
        #endregion

        #region 变量
        public string ProName;
        public string TheKeyword;
        public string NodeName;
        #endregion

        #region 初始化
        public CommonControl14()
         {
             InitializeComponent();
         }
        #endregion

        #region 窗体事件
         //查询
         private void toolStripButton2_Click(object sender, EventArgs e)
         {
            //YSGXCL 印刷工序产量
            //YHGXCL 印后工序产量
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
                    new SqlParameter("@DAH1",SqlDbType.VarChar),
                    new SqlParameter("@DAH2",SqlDbType.VarChar)
                   };
                    para[0].Value = toolStripTextBox1.Text;
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

        #region 方法

        #endregion


    }
}
