using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SqlConnect;
using System.Data.SqlClient;
using System.IO;


namespace ERPInquire
{
    public partial class Login : Form
    {
        #region 定量

         private static string strInformation = "在与数据库建立连接时出现与网络相关的错误——未找到或无法访问服务器。请检查本机是否可以连接上网络";
         private static string strInformation1 = "请核对用户名或密码";
         private static string strInformation2 = "提示";
         
        #endregion

        #region 变量
         private string sSkinPath = Application.StartupPath + @"\skin\皮肤\MacOS\";//获取皮肤的路径
         //初始化鼠标位置
         bool beginMove = false;
         int currentXPosition;
         int currentYPosition;
         
         public string userID;
         public string userName;
         public string auth;

        #endregion

        #region 初始化
        public Login()
        {
            InitializeComponent();
            //确定一开始登录窗体在屏幕的位置
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            SqlConnect.UserList s = new SqlConnect.UserList();
            loadsuser();
        }
        private void Login_Load(object sender, EventArgs e)
        {
            //comboBox2.SelectedIndex = 0;
            this.skinEngine1.SkinFile = sSkinPath + "MacOS.ssk";
            //登录自动填写用户名或密码
            ReadTXT();
            ReadTpwd();
        }
        #endregion

        #region 窗体事件

        //鼠标移动无边框窗体的事件
        private void Login_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                beginMove = true;
                currentXPosition = MousePosition.X;//鼠标的x坐标为当前窗体左上角x坐标
                currentYPosition = MousePosition.Y;//鼠标的y坐标为当前窗体左上角y坐标
            }
        }

        private void Login_MouseMove(object sender, MouseEventArgs e)
        {
            if (beginMove)
            {
                this.Left += MousePosition.X - currentXPosition;//根据鼠标x坐标确定窗体的左边坐标x
                this.Top += MousePosition.Y - currentYPosition;//根据鼠标的y坐标窗体的顶部，即Y坐标
                currentXPosition = MousePosition.X;
                currentYPosition = MousePosition.Y;
            }
        }

        private void Login_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                currentXPosition = 0; //设置初始状态
                currentYPosition = 0;
                beginMove = false;
            }
        }

        //登录按钮
        private void button1_Click(object sender, EventArgs e)
        {
            //核对登录密码
            checkpassword();
        }

        //清除按钮
        private void button2_Click(object sender, EventArgs e)
        {
            //清除密码并使密码框获得焦点
            textBox1.Text = "";
            textBox1.Focus();
        }

        //配置信息
        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            MenuBar.Authen authen = new MenuBar.Authen();
            authen.ShowDialog();
        }

        //关闭按钮
        private void button4_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }
        private void button4_MouseMove(object sender, MouseEventArgs e)
        {
            button4.BackColor = ColorTranslator.FromHtml("Red");
            button4.ForeColor = ColorTranslator.FromHtml("White");
        }
        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = ColorTranslator.FromHtml("White");
            button4.ForeColor = ColorTranslator.FromHtml("Black");
        }

        #endregion

        #region 方法
        private void ReadTXT()
        {
            if (!File.Exists("data/name.txt"))
            {
                //FileStream fs1 = new FileStream("data/name.txt", FileMode.Create, FileAccess.Write);//创建写入文件 
                //StreamWriter sw = new StreamWriter(fs1);
                //sw.WriteLine(this.textBox3.Text.Trim() + "+" + this.textBox4.Text);//开始写入值
                //sw.Close();
                //fs1.Close();
            }
            else
            {
                FileStream fs = new FileStream("data/name.txt", FileMode.Open, FileAccess.Read);
                StreamReader sd = new StreamReader(fs);
                //StreamWriter sr = new StreamWriter(fs);
                //sr.WriteLine(this.textBox3.Text.Trim() + "+" + this.textBox4.Text);//开始写入值
                comboBox1.Text = sd.ReadLine();
                fs.Close();
            }
        }
        private void ReadTpwd()
        {
            if (!File.Exists("data/pwd.txt"))
            {
                //FileStream fs1 = new FileStream("data/name.txt", FileMode.Create, FileAccess.Write);//创建写入文件 
                //StreamWriter sw = new StreamWriter(fs1);
                //sw.WriteLine(this.textBox3.Text.Trim() + "+" + this.textBox4.Text);//开始写入值
                //sw.Close();
                //fs1.Close();
            }
            else
            {
                FileStream fs = new FileStream("data/pwd.txt", FileMode.Open, FileAccess.Read);
                StreamReader sd = new StreamReader(fs);
                //StreamWriter sr = new StreamWriter(fs);
                //sr.WriteLine(this.textBox3.Text.Trim() + "+" + this.textBox4.Text);//开始写入值
                textBox1.Text = sd.ReadLine();
                fs.Close();
            }
        }
        private void WriteTXT()
        { 
                FileStream fs = new FileStream("data/name.txt", FileMode.Create, FileAccess.Write);
                StreamWriter sr = new StreamWriter(fs);
                sr.WriteLine(comboBox1.Text);//开始写入值
                sr.Close();
                fs.Close();
        }
        private void Writepwd()
        {
            FileStream fs = new FileStream("data/pwd.txt", FileMode.Create, FileAccess.Write);
            StreamWriter sr = new StreamWriter(fs);
            sr.WriteLine(textBox1.Text);//开始写入值
            sr.Close();
            fs.Close();
        }
        //初始化comboBox1控件，查找UserLogo文件中Active为1的人员，装载到下拉控件中
        private void loadsuser()
        {
            SqlConnect.UserList userlist = new SqlConnect.UserList();
            foreach (DataRow row in userlist.ds.Tables["UserLogo"].Select("ERP_Active=1"))
            {
                comboBox1.Items.Add(row["UserID"].ToString() + "-" + row["UserName"].ToString());
            }
        }

        //验证密码
        private void checkpassword()
        {
            //程序的超级用户权限，用于系统人员登录使用
            if (comboBox1.Text == "super")
            {
                if (textBox1.Text == "hyp")
                {
                    this.userID = "00000";
                    this.userName = "admin";
                    this.auth = "super";
                    MainFrm.userID = this.userID;
                    MainFrm.userName = this.userName;
                    MainFrm.auth = this.auth;
                    MainFrm.Dialog = "OK";
                    this.Close();
                }
                else { return; }
            }
            else if (comboBox1.Text == "" || textBox1.Text == "")
            {
                MessageBox.Show(strInformation1, strInformation2, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            }
            else
            {
                //查找登录密码
                try
                {
                    DataTable dt = new DataTable();
                    string[] a = comboBox1.Text.Split('-');
                    string sqls = "select password from ZCX_YHQX_M where id='" + a[0] + "' ";
                    SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
                    dt = connect.GetDataTable(sqls, "view");
                    SqlHelper.GetConnection().Close();

                    //查找用户装载界面权限
                    DataTable dt1 = new DataTable();
                    string sql1 = "select D.PermissionControl1,D.PermissionControl2 from ZCX_YHQX_M M left join ERPInquire_S_StaffTableM D on M.ERP_Authen = D.PermissionsID where M.id = '" + a[0] + "' ";
                    SqlConnect.ConnectStr connect1 = new SqlConnect.ConnectStr("ERP");
                    dt1 = connect.GetDataTable(sql1, "view");
                    SqlHelper.GetConnection().Close();

                    string per1 = Convert.ToString(dt1.Rows[0]["PermissionControl1"].ToString());
                    string per2 = Convert.ToString(dt1.Rows[0]["PermissionControl2"].ToString());

                    string pass = Convert.ToString(dt.Rows[0]["password"].ToString());

                    if (pass.Equals(textBox1.Text.Trim()) && a.Length == 2)
                    {
                        WriteTXT();
                        Writepwd();
                        this.userID = a[0];
                        this.userName = a[1];
                        //查找用户菜单栏的权限并返回一个DataSet
                        MainFrm.ds1 = SelectToolStripName(per1);
                        //查找用户树形目录的权限并返回一个DataSet
                        MainFrm.ds2 = SelectTreeName(per2);
                        MainFrm.userID = this.userID;
                        MainFrm.userName = this.userName;
                        MainFrm.auth = this.auth;
                        MainFrm.Dialog = "OK";
                        this.Close();
                    }
                    //
                    else
                    {
                        MessageBox.Show(strInformation1, strInformation2, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    }
                }
                catch(Exception ex)
                {
                     MessageBox.Show(ex.Message, strInformation2, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Warning);
                    //MessageBox.Show(ex.Message);
                }
            }
        }
        private DataSet SelectToolStripName(string str)
        {
            SqlHelper.GetConnection();
            DataSet ds = new DataSet();
            string[] condititons = str.Split(',');
            string ss = "select M.ModuleName AS 父模块名称,D.ModuleNameChild AS 子模块名称 from ERPInquire_S_MenuBarControlM M left join   ERPInquire_S_MenuBarControlD1 D  on  M.ID=D.FID where D.PermissionControl in (";
            for (int i = 0; i < condititons.Length; i++)
            {
                ss = ss + condititons[i] + ",";
            }
            ss = ss.Substring(0, ss.Length - 1);
            ss = ss + ") order by convert(int,M.ID), convert(int,D.ID) ";
            ds = SqlHelper.ExecuteDataSet(ss, null);
            SqlHelper.GetConnection().Close();
            return ds;
        }
        private DataSet SelectTreeName(string str)
        {
            SqlHelper.GetConnection();
            DataSet ds = new DataSet();
            string[] condititons = str.Split(',');
            string ss = "SELECT D1.ModuleName AS 树形模块名称, D.ModuleNameChild AS 子模块名称, M.ModuleNameChild AS 孙模块名称,M.StoredProcedureName AS 存储过程名称,M.ModuleCategory AS 模块类别,M.TheKeyword AS 关键字,M.KeyTime as 时间关键字, M.TheKeyword1 AS 关键字1 FROM ERPInquire_S_TreeMenuBarControlD1_D1 M LEFT JOIN ERPInquire_S_TreeMenuBarControlD1 D ON M.FID = D.ID left join ERPInquire_S_TreeMenuBarControlM D1 on m.GID = D1.ID WHERE  M.PermissionControl IN (";
            for (int i = 0; i < condititons.Length; i++)
            {
                ss = ss + condititons[i] + ",";
            }
            ss = ss.Substring(0, ss.Length - 1);
            ss = ss + ")   order by convert(int,D1.ID),CONVERT (INT,D.ID),M.TheOrder";
            ds = SqlHelper.ExecuteDataSet(ss, null);
            SqlHelper.GetConnection().Close();
            return ds;
        }

        #endregion

    } 
  }