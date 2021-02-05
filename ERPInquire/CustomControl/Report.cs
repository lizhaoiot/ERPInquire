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
    public partial class Report : UserControl
    {
        DateTimePicker sdate;
        DateTimePicker edate;
        TextBox filter ;
        public string sqlstr;
        public string modulename;
        public ConnectStr connect;
        DataTable dt = new DataTable();
        public Report()
        {
            InitializeComponent();       
            Initialize();
        }
        private void Initialize()
    {
            DateTime dt = DateTime.Today;
        this.sdate = new DateTimePicker();
        this.edate = new DateTimePicker();
        this.sdate.Value  = dt.AddDays(-2d);
        this.edate.Value = dt;
        AddToolstrip.AddDTPtoToolstrip(5, sdate, toolStrip1);
        AddToolstrip.AddDTPtoToolstrip(7, edate, toolStrip1);
    }
      
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
          ExcleIO.saveexcle(dt);
        }

     
        /****************/
 
        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void Report_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = modulename;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            sqlquery();
          //  exclequrey();
        }
        private void sqlquery()
        {

         string    sqls = sqlstr.Replace("@st", "'"+sdate.Value.ToString()+"'");
         sqls = sqls.Replace("@et", "'" + edate.Value.ToString() + "'");
         sqls = sqls.Replace("@key", "'" +'%'+ toolStripTextBox1.Text +'%'+ "'");
        //  MessageBox.Show(sqls);
         dt = connect.GetDataTable(sqls, "view");
         dataGridView1.DataSource = dt;
         dataGridView1.Refresh();
        }

        public static DataTable ExecuteStoredPro(string storeName, out int resultcount, string ID, string STARTDATE,string ENDDATE)
        {
            string connectionString = null;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = storeName;
                    SqlParameter[] para ={
                                     new SqlParameter("@ID",SqlDbType.VarChar),
                                     new SqlParameter("@STARTDATE",SqlDbType.DateTime),
                                     new SqlParameter("@ENDDATE",SqlDbType.DateTime),
                                     new SqlParameter("@result_value",SqlDbType.Int)

              };
                    para[0].Value = ID;
                    para[1].Value = STARTDATE;
                    para[2].Value = ENDDATE;
                    para[3].Direction = ParameterDirection.Output; //设定参数的输出方向  

                    cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    resultcount = Convert.ToInt32(cmd.Parameters[3].Value);
                    return dt;
                }
            }
        }
        /**********/
    }
}
