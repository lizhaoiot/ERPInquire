using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace ERPInquire.MenuBar
{
    public partial class ChangeMaterialJ : Form
    {


        #region 定量
          private static ChangeMaterialJ frm = null;
        #endregion

        #region 变量
          DataTable dt = new DataTable();
          //用户姓名
          public string Username = string.Empty;
          //用户部门
          public string Userbumen = string.Empty;
        #endregion



        #region 初始化
        private ChangeMaterialJ()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }
        public static ChangeMaterialJ CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new ChangeMaterialJ();
            }
            return frm;
        }
        private void ChangeMaterialJ_Load(object sender, EventArgs e)
        {
            textBox3.Text = "0";
            textBox5.Text = "0";
        }
        #endregion

        #region 窗体事件
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
                    cmd.CommandText = "thequery2";
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
                        dataGridView1.Columns["部件编号"].Visible = false;
                        dataGridView1.Columns["序号"].Visible = false;
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
            if (textBox2.Text.Trim().Equals(""))
            {
                return;
            }

            this.Cursor = Cursors.WaitCursor;//等待
            //判断纸张大小
            DataTable dtzl = SqlHelper.ExecuteDataTable("SELECT CD,KD FROM JCZL_WLBM_M WHERE WLBM='" + textBox2.Text.Trim() + "'");
            String CD = dtzl.Rows[0][0].ToString();
            String KD = dtzl.Rows[0][0].ToString();

            if (textBox6.Text.Trim().Equals(""))
            {
                try
                {
                    int index1 = 0;
                    int index2 = 0;
                    List<string> gdbh1 = new List<string>();
                    List<string> gdbh2 = new List<string>();
                    List<string> cpbh1 = new List<string>();
                    List<string> cpbh2 = new List<string>();
                    List<string> bjbh = new List<string>();
                    List<string> wlbh1 = new List<string>();
                    List<string> wlbh2 = new List<string>();
                    List<string> xh1 = new List<string>();
                    List<string> xh2 = new List<string>();
                    List<string> sl1 = new List<string>();
                    List<string> sl2 = new List<string>();
                    List<string> ks1 = new List<string>();
                    List<string> ks2 = new List<string>();
                    string SJC = string.Empty;
                    string SJK = string.Empty;

                    if (dataGridView1.Rows.Count > 0)
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells[0];
                            Boolean flag = Convert.ToBoolean(checkCell.Value);
                            if (flag)
                            {
                                if (!string.IsNullOrEmpty(dataGridView1.Rows[i].Cells["部件编号"].Value.ToString()))
                                {
                                    //工单产品部件物料
                                    gdbh1.Add(dataGridView1.Rows[i].Cells["生产施工单单号"].Value.ToString());
                                    cpbh1.Add(dataGridView1.Rows[i].Cells["产品编号"].Value.ToString());
                                    bjbh.Add(dataGridView1.Rows[i].Cells["部件编号"].Value.ToString());
                                    wlbh1.Add(dataGridView1.Rows[i].Cells["物料编号"].Value.ToString());
                                    xh1.Add(dataGridView1.Rows[i].Cells["序号"].Value.ToString());
                                    sl1.Add(dataGridView1.Rows[i].Cells["数量"].Value.ToString());
                                    ks1.Add(dataGridView1.Rows[i].Cells["开数"].Value.ToString());
                                    SJC = dataGridView1.Rows[i].Cells["上机长"].Value.ToString();
                                    SJK = dataGridView1.Rows[i].Cells["上机宽"].Value.ToString();
                                    index1 = index1 + 1;
                                }
                                else
                                {
                                    //工单产品物料
                                    gdbh2.Add(dataGridView1.Rows[i].Cells["生产施工单单号"].Value.ToString());
                                    cpbh2.Add(dataGridView1.Rows[i].Cells["产品编号"].Value.ToString());
                                    wlbh2.Add(dataGridView1.Rows[i].Cells["物料编号"].Value.ToString());
                                    xh2.Add(dataGridView1.Rows[i].Cells["序号"].Value.ToString());
                                    sl2.Add(dataGridView1.Rows[i].Cells["数量"].Value.ToString());
                                    ks2.Add(dataGridView1.Rows[i].Cells["开数"].Value.ToString());
                                    index2 = index2 + 1;
                                }
                            }
                        }
                        if (index1 > 1)
                        {
                            MessageBox.Show("一次只允许更新一种物料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        if (index2 > 1)
                        {
                            MessageBox.Show("一次只允许更新一种物料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        //更新数据
                        
                        for (int ij1 = 0; ij1 < index1; ij1++)
                        {
                            Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
                            SqlParameter[] sp ={
                                       new SqlParameter("@gdbh",SqlDbType.VarChar),
                                       new SqlParameter("@cpbh",SqlDbType.VarChar),
                                       new SqlParameter("@bjbh",SqlDbType.VarChar),
                                       new SqlParameter("@wlbh",SqlDbType.VarChar),
                                       new SqlParameter("@wlbhnew",SqlDbType.VarChar),
                                       new SqlParameter("@sl",SqlDbType.VarChar),
                                       new SqlParameter("@ks",SqlDbType.VarChar)
                        };
                            //判断纸张的开数
                            //获得原先纸张的面积
                            double prop = Convert.ToDouble((Convert.ToDouble(CD) * Convert.ToDouble(KD)) / (Convert.ToDouble(SJC) * Convert.ToDouble(SJK)));
                            sp[0].Value = gdbh1[ij1];
                            sp[1].Value = cpbh1[ij1];
                            sp[2].Value = bjbh[ij1];
                            sp[3].Value = wlbh1[ij1];
                            sp[4].Value = textBox2.Text;
                            sp[5].Value = textBox3.Text;
                            sp[6].Value = Math.Floor(prop).ToString();
                            Com.Hui.iMRP.Utils.SqlHelper.ExecStoredProcedure("theupdate3", sp);
                            Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
                            SqlHelper.ExecCommand("UPDATE SC_SCSGD_D2_D1_D2 SET REMARK='" + richTextBox1.Text + "' WHERE xh=" + xh1[ij1] + "");
                            string stemp = "生产施工单单号【" + gdbh1[ij1] + "】产品编码【" + cpbh1[ij1] + "】部件编码 【" + bjbh[ij1] + "】物料编码【" + wlbh1[ij1] + "】更改为 物料编码【" + textBox2.Text + "】";
                            Writelog(stemp);
                        }
                        

                       
                        for (int ij2 = 0; ij2 < index2; ij2++)
                        {
                            Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
                            SqlParameter[] sp ={
                                      new SqlParameter("@gdbh",SqlDbType.VarChar),
                                      new SqlParameter("@cpbh",SqlDbType.VarChar),
                                      new SqlParameter("@wlbh",SqlDbType.VarChar),
                                      new SqlParameter("@wlbhnew",SqlDbType.VarChar)
                        };
                            sp[0].Value = gdbh2[ij2];
                            sp[1].Value = cpbh2[ij2];
                            sp[2].Value = wlbh2[ij2];
                            sp[3].Value = textBox2.Text;
                            Com.Hui.iMRP.Utils.SqlHelper.ExecStoredProcedure("theupdate4", sp);
                            Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
                            SqlHelper.ExecCommand("UPDATE SC_SCSGD_D2_D3 SET REMARK='" + richTextBox1.Text + "' WHERE xh=" + xh2[ij2] + "");
                            string stemp = "生产施工单单号【" + gdbh2[ij2] + "】产品编码【" + cpbh2[ij2] + "】物料编码【" + wlbh2[ij2] + "】更改为 物料编码【" + textBox2.Text + "】";
                            Writelog(stemp);
                        }
                      

                        MessageBox.Show("更新成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            if (!textBox6.Text.Trim().Equals(""))
            {
                try
                {
                    int index1 = 0;
                    int index2 = 0;
                    List<string> gdbh1 = new List<string>();
                    List<string> gdbh2 = new List<string>();
                    List<string> cpbh1 = new List<string>();
                    List<string> cpbh2 = new List<string>();
                    List<string> bjbh = new List<string>();
                    List<string> wlbh1 = new List<string>();
                    List<string> wlbh2 = new List<string>();
                    List<string> xh1 = new List<string>();
                    List<string> xh2 = new List<string>();
                    List<string> sl1 = new List<string>();
                    List<string> sl2 = new List<string>();
                    List<string> ks1 = new List<string>();
                    List<string> ks2 = new List<string>();
                    string SJC = string.Empty;
                    string SJK = string.Empty;

                    if (dataGridView1.Rows.Count > 0)
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells[0];
                            Boolean flag = Convert.ToBoolean(checkCell.Value);
                            if (flag)
                            {
                                if (!string.IsNullOrEmpty(dataGridView1.Rows[i].Cells["部件编号"].Value.ToString()))
                                {
                                    //工单产品部件物料
                                    gdbh1.Add(dataGridView1.Rows[i].Cells["生产施工单单号"].Value.ToString());
                                    cpbh1.Add(dataGridView1.Rows[i].Cells["产品编号"].Value.ToString());
                                    bjbh.Add(dataGridView1.Rows[i].Cells["部件编号"].Value.ToString());
                                    wlbh1.Add(dataGridView1.Rows[i].Cells["物料编号"].Value.ToString());
                                    xh1.Add(dataGridView1.Rows[i].Cells["序号"].Value.ToString());
                                    sl1.Add(dataGridView1.Rows[i].Cells["数量"].Value.ToString());
                                    ks1.Add(dataGridView1.Rows[i].Cells["开数"].Value.ToString());
                                    SJC = dataGridView1.Rows[i].Cells["上机长"].Value.ToString();
                                    SJK = dataGridView1.Rows[i].Cells["上机宽"].Value.ToString();
                                    index1 = index1 + 1;
                                }
                                else
                                {
                                    //工单产品物料
                                    gdbh2.Add(dataGridView1.Rows[i].Cells["生产施工单单号"].Value.ToString());
                                    cpbh2.Add(dataGridView1.Rows[i].Cells["产品编号"].Value.ToString());
                                    wlbh2.Add(dataGridView1.Rows[i].Cells["物料编号"].Value.ToString());
                                    xh2.Add(dataGridView1.Rows[i].Cells["序号"].Value.ToString());
                                    sl2.Add(dataGridView1.Rows[i].Cells["数量"].Value.ToString());
                                    ks2.Add(dataGridView1.Rows[i].Cells["开数"].Value.ToString());
                                    index2 = index2 + 1;
                                }
                            }
                        }
                        if (index1 > 1)
                        {
                            MessageBox.Show("一次只允许更新一种物料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        if (index2 > 1)
                        {
                            MessageBox.Show("一次只允许更新一种物料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        //更新数据
                      
                        for (int ij1 = 0; ij1 < index1; ij1++)
                        {
                            Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
                            SqlParameter[] sp ={
                                       new SqlParameter("@gdbh",SqlDbType.VarChar),
                                       new SqlParameter("@cpbh",SqlDbType.VarChar),
                                       new SqlParameter("@bjbh",SqlDbType.VarChar),
                                       new SqlParameter("@wlbh",SqlDbType.VarChar),
                                       new SqlParameter("@wlbhnew",SqlDbType.VarChar),
                                       new SqlParameter("@sl",SqlDbType.VarChar),
                                       new SqlParameter("@ks",SqlDbType.VarChar)
                        };
                            //判断纸张的开数
                            //获得原先纸张的面积
                            double prop = Convert.ToDouble((Convert.ToDouble(CD) * Convert.ToDouble(KD)) / (Convert.ToDouble(SJC) * Convert.ToDouble(SJK)));
                            sp[0].Value = gdbh1[ij1];
                            sp[1].Value = cpbh1[ij1];
                            sp[2].Value = bjbh[ij1];
                            sp[3].Value = wlbh1[ij1];
                            sp[4].Value = textBox2.Text;
                            sp[5].Value = textBox3.Text;
                            sp[6].Value = Math.Floor(prop).ToString();
                            Com.Hui.iMRP.Utils.SqlHelper.ExecStoredProcedure("theupdate3", sp);
                            Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
                            SqlHelper.ExecCommand("UPDATE SC_SCSGD_D2_D1_D2 SET REMARK='" + richTextBox1.Text + "' WHERE xh=" + xh1[ij1] + "");
                            string stemp = "生产施工单单号【" + gdbh1[ij1] + "】产品编码【" + cpbh1[ij1] + "】部件编码 【" + bjbh[ij1] + "】物料编码【" + wlbh1[ij1] + "】更改为 物料编码【" + textBox2.Text + "】";
                            Writelog(stemp);
                        }
                        

                        
                        for (int ij2 = 0; ij2 < index2; ij2++)
                        {
                            Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
                            SqlParameter[] sp ={
                                      new SqlParameter("@gdbh",SqlDbType.VarChar),
                                      new SqlParameter("@cpbh",SqlDbType.VarChar),
                                      new SqlParameter("@wlbh",SqlDbType.VarChar),
                                      new SqlParameter("@wlbhnew",SqlDbType.VarChar)
                        };
                            sp[0].Value = gdbh2[ij2];
                            sp[1].Value = cpbh2[ij2];
                            sp[2].Value = wlbh2[ij2];
                            sp[3].Value = textBox2.Text;
                            Com.Hui.iMRP.Utils.SqlHelper.ExecStoredProcedure("theupdate4", sp);
                            Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
                            SqlHelper.ExecCommand("UPDATE SC_SCSGD_D2_D3 SET REMARK='" + richTextBox1.Text + "' WHERE xh=" + xh2[ij2] + "");
                            string stemp = "生产施工单单号【" + gdbh2[ij2] + "】产品编码【" + cpbh2[ij2] + "】物料编码【" + wlbh2[ij2] + "】更改为 物料编码【" + textBox2.Text + "】";
                            Writelog(stemp);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
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
        //按照工单号查询
        private void button5_Click(object sender, EventArgs e)
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
                    cmd.CommandText = "thequery4";
                    SqlParameter[] para ={
                                     new SqlParameter("@DAH",SqlDbType.VarChar)
                    };
                    para[0].Value = textBox4.Text.Trim();
                    try
                    {
                        cmd.CommandTimeout = 60 * 60 * 1000;
                        cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dt);
                        conn.Close();
                        dataGridView1.DataSource = dt;
                        dataGridView1.Columns["部件编号"].Visible = false;
                        dataGridView1.Columns["序号"].Visible = false;
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
        #endregion
        //查询

        private void Writelog(string s)
        {
            Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
            String S = "insert into ERPInquireLog (usersloadname,mac,ip,operatingtime,bumen,describe,machinename,usermachinename) values('" + Username + "','" + GetMacAddress() + "','" + GetClientLocalIPv4Address() + "','" + GetUserstime() + "','" + Userbumen + "','" + s + "','" + GetMachineName() + "','" + GetUsersMachineName() + "')";
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
            return ipAddr.ToString();
        }
        private string GetMachineName()
        {
            return System.Environment.MachineName;
        }
        private string GetUsersMachineName()
        {
            return System.Environment.UserName;
        }
        private string GetUserstime()
        {
            return System.DateTime.Now.ToString();
        }

    }
}
