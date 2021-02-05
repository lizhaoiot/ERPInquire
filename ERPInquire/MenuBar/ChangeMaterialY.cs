using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Com.Hui.Controls;
using System.Data.SqlClient;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.IO;

namespace ERPInquire.MenuBar
{
    public partial class ChangeMaterialY : Form
    {
        #region 定量
        private static ChangeMaterialY frm = null;
        #endregion

        #region 变量
        DataTable dt = new DataTable();
        //用户姓名
        public  string Username = string.Empty;
        //用户部门
        public  string Userbumen = string.Empty;
        #endregion

        #region 初始化
        private ChangeMaterialY()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }
        public static ChangeMaterialY CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new ChangeMaterialY();
            }
            return frm;
        }
        private void ChangeMaterialY_Load(object sender, EventArgs e)
        {

        }
        #endregion

        #region 窗体事件
        //查询
        private void button1_Click(object sender, EventArgs e)
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
                        cmd.CommandText = "thequery1";
                        SqlParameter[] para ={
                                     new SqlParameter("@DAH",SqlDbType.VarChar)
                    };
                    para[0].Value = textBox1.Text.Trim();
                    try
                     {
                         cmd.CommandTimeout = 60 * 60 * 1000;
                         cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                         SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                         adapter.Fill(dt);
                         conn.Close();           
                         dataGridView1.DataSource = dt;
                         dataGridView1.Columns["部件编码"].Visible = false;
                         dataGridView1.RowsDefaultCellStyle.BackColor = Color.Azure;
                         dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
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
        //更改
        private void button2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            try
            {
                int index1 = 0;
                int index2 = 0;
                List<string> cpbh1 = new List<string>();//产品编码
                List<string> cpbh2 = new List<string>();//产品编码
                List<string> bjbh = new List<string>();//部件编码
                List<string> wlbh1 = new List<string>();//物料编码
                List<string> wlbh2 = new List<string>();//物料编码
                if (dataGridView1.Rows.Count > 0)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells[0];
                        Boolean flag = Convert.ToBoolean(checkCell.Value);
                        if (flag)
                        {
                            if (!string.IsNullOrEmpty(dataGridView1.Rows[i].Cells["部件编码"].Value.ToString()))
                            {
                               //产品部件物料
                                cpbh1.Add(dataGridView1.Rows[i].Cells["产品编码"].Value.ToString());
                                bjbh.Add(dataGridView1.Rows[i].Cells["部件编码"].Value.ToString());
                                wlbh1.Add(dataGridView1.Rows[i].Cells["物料编码"].Value.ToString());
                                index1 = index1 + 1;
                            }
                            else
                            {
                                //产品物料
                                cpbh2.Add(dataGridView1.Rows[i].Cells["产品编码"].Value.ToString());
                                wlbh2.Add(dataGridView1.Rows[i].Cells["物料编码"].Value.ToString());
                                index2 = index2 + 1;
                            }
                        }
                    }  
                    //更新数据
                    #region 更新产品部件物料数据
                for (int ij1 = 0; ij1 < index1; ij1++)
                {
                    Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
                    SqlParameter[] sp ={
                                     new SqlParameter("@cpbh",SqlDbType.VarChar),
                                     new SqlParameter("@bjbh",SqlDbType.VarChar),
                                     new SqlParameter("@wlbh",SqlDbType.VarChar),
                                     new SqlParameter("@wlbhnew",SqlDbType.VarChar)
                    };
                    sp[0].Value = cpbh1[ij1];
                    sp[1].Value = bjbh[ij1];
                    sp[2].Value = wlbh1[ij1];
                    sp[3].Value = textBox2.Text;
                    Com.Hui.iMRP.Utils.SqlHelper.ExecStoredProcedure("theupdate1", sp);
                    Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
                    string stemp = "产品编码【" + cpbh1[ij1] + "】部件编码【" + bjbh[ij1] + "】物料编码【" + wlbh1[ij1] + "】更改为 物料编码【" + textBox2.Text + "】";
                    Writelog(stemp);
                }
                #endregion
                    
                    #region 更新产品物料数据
                for (int ij2 = 0; ij2 < index2; ij2++)
                {
                    Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
                    SqlParameter[] sp ={
                                     new SqlParameter("@cpbh",SqlDbType.VarChar),
                                     new SqlParameter("@wlbh",SqlDbType.VarChar),
                                     new SqlParameter("@wlbhnew",SqlDbType.VarChar)
                    };
                    sp[0].Value = cpbh2[ij2];
                    sp[1].Value = wlbh2[ij2];
                    sp[2].Value = textBox2.Text;
                    Com.Hui.iMRP.Utils.SqlHelper.ExecStoredProcedure("theupdate2", sp);
                    Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
                    string stemp = "产品编码【" + cpbh2[ij2] + "】物料编码【" + wlbh2[ij2] + "】更改为 物料编码【" + textBox2.Text + "】";
                    Writelog(stemp);
                }
                #endregion
                                  
                    MessageBox.Show("更新成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            this.Cursor = Cursors.Default;//正常状态
        }
        //全选
        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = true;
            }
        }
        //全不选
        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = false;
            }
        }
        #endregion

        #region 方法
        //数据库命令行操作
        
        private void Writelog(string s)
        {
            Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
            String S = "insert into ERPInquireLog(usersloadname,mac,ip,operatingtime,bumen,describe,numbers,machinename,usermachinename) values('" + Username + "','" + GetMacAddress() + "','" + GetClientLocalIPv4Address() + "','" + GetUserstime() + "','" + Userbumen + "','" + s + "','" + GetMachineName() + "','" + GetUsersMachineName() + "')";
            Com.Hui.iMRP.Utils.SqlHelper.ExecCommand(S);
            Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
        }
        //得到MAC地址
        private string GetMacAddress()
        {
            try
            {
                string strMac = string.Empty;
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    if ((bool)mo["IPEnabled"] == true)
                    {
                        strMac = mo["MacAddress"].ToString();
                    }
                }
                moc = null;
                mc = null;
                return strMac;
            }
            catch
            {
                return "unknown";
            }
        }
        //得到IPV4
        private string GetClientLocalIPv4Address()
        {
            IPAddress ipAddr = Dns.Resolve(Dns.GetHostName()).AddressList[0];//获得当前IP地址
            return  ipAddr.ToString();
        }
        private  string GetMachineName()
        {
           return  System.Environment.MachineName;   
        }
        private string GetUsersMachineName()
        {
            return System.Environment.UserName;   
        }
        private string GetUserstime()
        {
            return System.DateTime.Now.ToString();
        }
        #endregion
    }
}
