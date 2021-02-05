using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SqlConnect;
using System.Data.SqlClient;
using ExcelOperation;

namespace ERPInquire.CustomControl
{
    public partial class CommonControl5 : UserControl
    {

        #region 变量

        public string ProName;
        public string NodeName;

        #endregion

        #region 定量

        DataTable dt = new DataTable();
        DataTable dtNew = new DataTable();

        #endregion

        #region 初始化 
        public CommonControl5()
        {
            InitializeComponent();
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行被选中
        }
        private void CommonControl5_Load(object sender, EventArgs e)
            {
                toolStripComboBox1.Text = DateTime.Now.Month.ToString();
                toolStripComboBox2.Text = DateTime.Now.Year.ToString();
            }
        #endregion

        #region 窗体事件

        //查询
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (toolStripComboBox1.Text == "") return;
            if (toolStripComboBox2.Text == "") return;
            if (toolStripTextBox1.Text.Equals(""))
            {
                dt.Clear();
                this.Cursor = Cursors.WaitCursor;//等待
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
                                     new SqlParameter("@Years",SqlDbType.Int),
                                     new SqlParameter("@Months",SqlDbType.Int),
                                     new SqlParameter("@Days",SqlDbType.Int)
                };
                        para[0].Value = Convert.ToInt16(toolStripComboBox2.Text);
                        para[1].Value = Convert.ToInt16(toolStripComboBox1.Text);
                        para[2].Value = Convert.ToInt16(PanDuan(toolStripComboBox2.Text, toolStripComboBox1.Text));
                        try
                        {
                            cmd.CommandTimeout = 60 * 60 * 1000;
                            cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(dt);
                            conn.Close();
                            DataGridViewShow(dt);
                            //获得datatable的列名
                            List<string> ls = SqlHelper.GetColumnsByDataTable(dt);
                            toolStripComboBox3.ComboBox.DataSource = ls;
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
                int i = toolStripComboBox3.ComboBox.SelectedIndex;
                string dtcolounname = toolStripComboBox3.SelectedItem.ToString();
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
        private void DataGridViewShow(DataTable d1)
        {
            dataGridView1.DataSource = d1;
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
            toolStripLabel3.Text = "总行数:" + (d1.Rows.Count).ToString();
            dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView1.Refresh();
        }
        //导出EXCEL
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
          MyMIS.ExcleIO.Export2Excel(NodeName, dt);
        }
        #endregion

        #region 判断年份
        //判断是不是闰年
        private bool RunNian(int year)
        {
            if ((year % 400 == 0) || (year % 4 == 0 && year % 100 != 0))
                return true;
            else
                return false;
        }
        //判断有多少天
       private int PanDuan(string years,string months)
        {
            int days = 0 ;
            int year = Convert.ToInt32(years);
            int month = Convert.ToInt32(months);
             switch (month)
             {
                 case 1:
                 case 3:
                 case 5:
                 case 7:
                 case 8:
                 case 10:
                 case 12:
                     days= 31;
                     break;
                 case 4:
                 case 6:
                 case 9:
                 case 11:
                     days = 30;
                     break;
                 case 2:
                     if (!(RunNian(Convert.ToInt32(years))))
                     {
                         days = 28;
                     }
                     else
                     {
                         days = 29;
                     }
                     break;
                 default:
                     break;
             }
            return days;
        }
        #endregion

        #region datatable行列转置

        #endregion

        #region 数据分析
        private void toolStripButton4_Click(object sender, EventArgs e)
        {

        }
        #endregion

    }
}
