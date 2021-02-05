using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SqlConnect;
using System.IO;
using System.Data.SqlClient;

namespace MyMIS
{
    public partial class SCSGD2 : UserControl
    {
        #region 定量

        #endregion

        #region 变量

        DateTimePicker sdate;
        DateTimePicker edate;
        TextBox filter;
        public string sqlstr;
        public string modulename;
        public ConnectStr connect;
        DataTable dt = new DataTable();
        string id;
        DateTime starttime;
        DateTime endtime;
        #endregion

        #region 初始化
        public SCSGD2()
        {
            InitializeComponent();
            Initialize();
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
            AddToolstrip.AddDTPtoToolstrip(5, sdate, toolStrip1);
            AddToolstrip.AddDTPtoToolstrip(7, edate, toolStrip1);
        }
        #endregion

        #region 窗体事件
        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            ExcleIO.saveexcle(dt);
        }

        private void SGSGD2_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = modulename;
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            id = toolStripTextBox1.Text;
            dataGridView1.DataSource = ExecuteStoredPro("produce5", id, starttime, endtime);
            dataGridView1.Refresh();
        }
        #endregion

        #region 方法
        public static DataTable ExecuteStoredPro(string storeName, string ID, DateTime STARTDATE, DateTime ENDDATE)
        {
            string connectionString = "data source=192.168.0.97; Database=hy;user id=sa; password=";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = storeName;
                    SqlParameter[] para ={
                                     new SqlParameter("@DHA",SqlDbType.VarChar),
                                     new SqlParameter("@STARTTIME",SqlDbType.DateTime),
                                     new SqlParameter("@ENDDATE",SqlDbType.DateTime)
               };
                    para[0].Value = ID;
                    para[1].Value = STARTDATE;
                    para[2].Value = ENDDATE;
                    cmd.CommandTimeout = 60 * 60 * 1000;
                    cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    conn.Close();
                    return dt;
                }
            }
        }
        #endregion

    }
}
