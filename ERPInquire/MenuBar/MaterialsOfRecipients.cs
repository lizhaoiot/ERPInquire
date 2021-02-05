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

namespace ERPInquire.MenuBar
{
    public partial class MaterialsOfRecipients : Form
    {

        #region 定量

        #endregion

        #region 变量
           private static string values = string.Empty;
           DateTimePicker sdate;
           DateTimePicker edate;
           public string sqlstr;
           public ConnectStr connect;
           DataTable dt = new DataTable();
           private static MaterialsOfRecipients frm = null;
        #endregion

        #region 初始化
          private MaterialsOfRecipients()
          {
              this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
              InitializeComponent();
              Initialize();
          }
          public static MaterialsOfRecipients CreateInstrance()
          {
              if (frm == null || frm.IsDisposed)
              {
                  frm = new MaterialsOfRecipients();
              }
              return frm;
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

        private void MaterialsOfRecipients_Load(object sender, EventArgs e)
        {
           //  this.TopMost = true;
        }
        #endregion

        #region 窗体事件
        //查询
        private void toolStripButton3_Click(object sender, EventArgs e)
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
                    cmd.CommandText = "YUAN_CLCRK_Query";
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
                        dataGridView1.RowsDefaultCellStyle.BackColor = Color.Azure;
                        dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
                        toolStripLabel2.Text = "总行数:" + (dt.Rows.Count).ToString();
                        dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
                        dataGridView1.Refresh();
                        this.dataGridView1.Columns["XH"].Visible = false;
                        this.dataGridView1.Columns["ZBXH"].Visible = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            this.Cursor = Cursors.Default;//正常状态
        }
        //打印
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            string s = string.Empty;
            SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
            string s1 = connect.GetConnectStr("ERP");
            int ijk = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Selected == true)
                {
                    s = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    ijk = ijk + 1;
                }
            }
            if (ijk > 1)
            {
                MessageBox.Show("请逐一单据打印", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            List<Com.Hui.Controls.ReportParameter> list = new List<Com.Hui.Controls.ReportParameter>();
            list.Add(new Com.Hui.Controls.ReportParameter("bdbh",s));
            Com.Hui.Controls.ReportHelper.PrintReport("cllydjyt.frx", list,s1,true);
        }
        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.Text == "编辑")
            {
                List<string> slist1 = new List<string>();
                List<string> slist2 = new List<string>();
                List<string> slist3 = new List<string>();
                List<string> slist4 = new List<string>();
                for (int i=0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Selected == true)
                    {
                       slist4.Add(dataGridView1.Rows[i].Cells[25].Value.ToString());
                       slist1.Add(dataGridView1.Rows[i].Cells[26].Value.ToString());
                    }
                }
                ERPInquire.MenuBar.MaterialRequirements mr = ERPInquire.MenuBar.MaterialRequirements.CreateInstrance();
                mr.st4 = slist4;
                mr.st1 = slist1;
                mr.Show();
            }
        }
        public void s1()
        {
            ERPInquire.MenuBar.MaterialRequirements mr = ERPInquire.MenuBar.MaterialRequirements.CreateInstrance();
            values = mr.ReturnValue();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Selected == true)
                {
                    dataGridView1.Rows[i].Cells[0].Value = values;
                }
            }
        }
        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Point contextMenuPoint = dataGridView1.PointToClient(Control.MousePosition);
                contextMenuStrip1.Show(dataGridView1, contextMenuPoint);
            }
        }

        #endregion

        #region 公共方法

        #endregion
    }
}
