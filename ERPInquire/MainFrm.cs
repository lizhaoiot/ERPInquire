using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERPInquire
{
    public partial class MainFrm : Form
    {
        #region 变量

        private SqlConnect.ConnectStr connect;
        public static string userID;
        public static string userName;
        public static string auth;
        public static string Dialog = null;
        public static DataSet ds1;
        public static DataSet ds2;
        private DataSet ds3 = new DataSet();
        private DataSet ds4 = new DataSet();
        private string[] str = null;
        private string[] strName1;
        private List<string> listCombobox;
        public TreeNode IsselectNode;
        private string sSkinPath = Application.StartupPath + @"\skin\MacOS\";//获取皮肤的路径
        #endregion 变量

        #region 工序
        //SQL SERVER
        DataColumn d10 = new DataColumn("产品编号");
        DataColumn d11 = new DataColumn("部件编号");
        DataColumn d12 = new DataColumn("部件名称");
        DataColumn d13 = new DataColumn("工序编码");
        DataColumn d14 = new DataColumn("工序类别");
        DataColumn d15 = new DataColumn("工序名称");
        //Sqlite
        DataColumn d20 = new DataColumn("CPBH");
        DataColumn d21 = new DataColumn("ZBXH_BJBH");
        DataColumn d22 = new DataColumn("BJMC");
        DataColumn d23 = new DataColumn("GXBH");
        DataColumn d24 = new DataColumn("GBLB");
        DataColumn d25 = new DataColumn("GXMC");
        #endregion

        #region 物料
        //SQL SERVER
        DataColumn d30 = new DataColumn("产品编号");
        DataColumn d31 = new DataColumn("产品名称");
        DataColumn d32 = new DataColumn("部件名称");
        DataColumn d33 = new DataColumn("物料大类");
        DataColumn d34 = new DataColumn("物料小类");
        DataColumn d35 = new DataColumn("物料编号");
        DataColumn d36 = new DataColumn("物料名称");
        //Sqlite
        DataColumn d40 = new DataColumn("CPBH");
        DataColumn d41 = new DataColumn("CPMC");
        DataColumn d42 = new DataColumn("BJMC");
        DataColumn d43 = new DataColumn("WLDL");
        DataColumn d44 = new DataColumn("WLZL");
        DataColumn d45 = new DataColumn("WLBH");
        DataColumn d46 = new DataColumn("WLMC");
        #endregion

        #region 初始化

        public MainFrm()
        {
            Login fl = new Login();
            fl.ShowDialog();
            if (Dialog == "OK")
            {
                InitializeComponent();
                //使窗体最大化并位于整个屏幕的中心
                //this.WindowState = FormWindowState.Maximized;
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

                toolStripStatusLabel1.Text = "当前登录用户：" + userName;

                connect = new SqlConnect.ConnectStr("ERP");

                this.toolStripStatusLabel3.Alignment = ToolStripItemAlignment.Right;
                if (userID.Equals("00000"))
                {
                    //装载全部界面
                    LoadToolStrip2(ds1);
                    LoadTree2(ds2);
                }
                else
                {
                    //根据用户权限装载部分界面
                    LoadToolStrip2(ds1);
                    LoadTree2(ds2);
                }
            }
        }

        private void MainFrm_Load(object sender, EventArgs e)
        {
            this.skinEngine1.SkinFile = sSkinPath + "MacOS.ssk";
            string[] strName = new string[ds2.Tables[0].Rows.Count];
            for (int i = 0; i < ds2.Tables[0].Rows.Count; i++)
            {
                strName[i] = ds2.Tables[0].Rows[i]["孙模块名称"].ToString();
            }
            strName1 = strName;
            comboBox1.Items.AddRange(strName1);
            listCombobox = getComboboxItems(this.comboBox1);//获取Item

            DataTable dt = ds2.Tables[0];
            List<string[]> list = new List<string[]>();
            foreach (DataRow r in dt.Rows)
            {
                int colCount = r.ItemArray.Count();
                string[] items = new string[colCount];
                for (int i = 0; i < colCount; i++)
                {
                    items[i] = Convert.ToString(r.ItemArray[i]);
                }
                list.Add(items);
            }

            this.panel1.Controls.Clear();
            CustomControl.Hello hello = new CustomControl.Hello();
            hello.Dock = DockStyle.Fill;
            hello.name = userName;
            this.panel1.Controls.Add(hello);
            ExpandTree2();

        }
        #endregion 初始化

        #region 窗体事件
        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, panel1.ClientRectangle,
            Color.Black, 1, ButtonBorderStyle.Solid, //左边
　　     　  Color.Black, 1, ButtonBorderStyle.Solid, //上边
　　　       Color.Black, 1, ButtonBorderStyle.Solid, //右边
　　         Color.Black, 1, ButtonBorderStyle.Solid);//底边
        }
        //装载按钮
        private void button1_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel2.Text = "模块名称：" + comboBox1.Text;
            this.panel1.Controls.Clear();
            DataTable table = ds2.Tables["TABLE"];
            string expression;
            expression = "孙模块名称 ='" + comboBox1.Text + "'";
            DataRow[] foundRows;
            //使用选择方法来找到匹配的所有行。
            foundRows = table.Select(expression);
            if (foundRows.Length == 1)
            {
                //过滤行,找到所要的行。
                string strTemp1 = foundRows[0]["存储过程名称"].ToString();
                string strTemp2 = foundRows[0]["模块类别"].ToString();
                string strTemp3 = foundRows[0]["关键字"].ToString();
                string strTemp4 = foundRows[0]["时间关键字"].ToString();
                string strTemp5 = foundRows[0]["关键字1"].ToString();

                if (strTemp1 != "" && strTemp2 != "")
                {
                    #region 前后日期以及关键字查询

                    if (strTemp2.Equals("1"))
                    {
                        CustomControl.CommonControl2 cmd2 = new CustomControl.CommonControl2();
                        cmd2.ProName = strTemp1;
                        cmd2.TheKeyword = strTemp3;
                        cmd2.TimeKey = strTemp4;
                        cmd2.Dock = DockStyle.Fill;
                        cmd2.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd2);
                    }

                    #endregion 前后日期以及关键字查询

                    #region 关键字查询

                    if (strTemp2.Equals("2"))
                    {
                        CustomControl.CommonControl1 cmd1 = new CustomControl.CommonControl1();
                        cmd1.ProName = strTemp1;
                        cmd1.TheKeyword = strTemp3;
                        cmd1.Dock = DockStyle.Fill;
                        cmd1.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd1);
                    }

                    #endregion 关键字查询

                    #region 批量查询

                    if (strTemp2.Equals("3"))
                    {
                        try
                        {
                            SqlHelper.GetConnection();
                            DataTable dt = new DataTable();
                            string s1 = "select * from ERPInquire_M_BatchQueryData where ModuleNameChild='" + comboBox1.Text + "'";
                            dt = SqlHelper.ExecuteDataTable(s1);
                            SqlHelper.GetConnection().Close();
                            if (dt.Rows.Count == 1)
                            {
                                string nameTemp = dt.Rows[0][2].ToString();
                                string arributeTemp = dt.Rows[0][3].ToString();
                                string[] nameTemp1 = nameTemp.Split(',');
                                string[] arributeTemp1 = arributeTemp.Split(',');
                                CustomControl.CommonControl3 cmd3 = new CustomControl.CommonControl3();
                                cmd3.names = nameTemp1;
                                cmd3.attributes = arributeTemp1;
                                cmd3.nodename = comboBox1.Text;
                                cmd3.ProName = strTemp1;
                                cmd3.NodeName = comboBox1.Text;
                                cmd3.Dock = DockStyle.Fill;
                                this.panel1.Controls.Add(cmd3);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }

                    #endregion 批量查询

                    #region 区间日期查询

                    if (strTemp2.Equals("4"))
                    {
                        CustomControl.CommonControl4 cmd4 = new CustomControl.CommonControl4();
                        cmd4.ProName = strTemp1;
                        cmd4.Dock = DockStyle.Fill;
                        cmd4.TimeKey = strTemp4;
                        cmd4.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd4);
                    }

                    #endregion 区间日期查询

                    #region 按月份统计

                    if (strTemp2.Equals("5"))
                    {
                        CustomControl.CommonControl5 cmd5 = new CustomControl.CommonControl5();
                        cmd5.ProName = strTemp1;
                        cmd5.Dock = DockStyle.Fill;
                        cmd5.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd5);
                    }

                    #endregion 按月份统计

                    #region 时间区间以及两个查询关键字
                    if (strTemp2.Equals("8"))
                    {
                        CustomControl.CommonControl8 cmd8 = new CustomControl.CommonControl8();
                        cmd8.ProName = strTemp1;
                        cmd8.TheKeyword = strTemp3;
                        cmd8.TheKeyword1 = strTemp5;
                        cmd8.TimeKey = strTemp4;
                        cmd8.Dock = DockStyle.Fill;
                        cmd8.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd8);
                    }
                    #endregion

                    #region 两个关键字
                    if (strTemp2.Equals("9"))
                    {
                        CustomControl.CommonControl9 cmd9 = new CustomControl.CommonControl9();
                        cmd9.ProName = strTemp1;
                        cmd9.TheKeyword = strTemp3;
                        cmd9.TheKeyword1 = strTemp5;
                        cmd9.Dock = DockStyle.Fill;
                        cmd9.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd9);
                    }
                    #endregion
                }
                if (strTemp1 == "" && strTemp2 != "")
                {
                    if (strTemp2.Equals("6"))
                    {
                        CustomControl.CommonControl6 cmd6 = new CustomControl.CommonControl6();
                        cmd6.Dock = DockStyle.Fill;
                        cmd6.TimeKey = strTemp4;
                        cmd6.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd6);
                    }
                    if (strTemp2.Equals("7"))
                    {
                        CustomControl.CommonControl7 cmd7 = new CustomControl.CommonControl7();
                        cmd7.Dock = DockStyle.Fill;
                        cmd7.TimeKey = strTemp4;
                        cmd7.NodeName = comboBox1.Text;
                        this.panel1.Controls.Add(cmd7);
                    }
                }

            }
            else
            {
                return;
            }
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            List<string> listNew = new List<string>();
            this.comboBox1.Items.Clear();
            listNew.Clear();
            foreach (string item in listCombobox)
            {
                if (item.Contains(this.comboBox1.Text))
                {
                    listNew.Add(item);
                }
            }
            if (listNew.Count != 0)
            {
                this.comboBox1.Items.AddRange(listNew.ToArray());
                this.comboBox1.SelectionStart = this.comboBox1.Text.Length;
                Cursor = Cursors.Default;
                this.comboBox1.DroppedDown = true;
            }
            else
            {
                this.comboBox1.Items.Add("");
                this.comboBox1.SelectionStart = this.comboBox1.Text.Length;
            }
        }

        //得到Combobox的数据，返回一个List
        public List<string> getComboboxItems(ComboBox cb)
        {
            //初始化绑定默认关键词
            List<string> listOnit = new List<string>();
            //将数据项添加到listOnit中
            for (int i = 0; i < cb.Items.Count; i++)
            {
                listOnit.Add(cb.Items[i].ToString());
            }
            return listOnit;
        }

        //模糊查询Combobox
        public void selectCombobox(ComboBox cb, List<string> listOnit)
        {
            //输入key之后返回的关键词
            List<string> listNew = new List<string>();
            //清空combobox
            cb.Items.Clear();
            //清空listNew
            listNew.Clear();
            //遍历全部备查数据
            foreach (var item in listOnit)
            {
                if (item.Contains(cb.Text))
                {
                  //符合，插入ListNew
                  listNew.Add(item);
                }
            }
            //combobox添加已经查询到的关键字
            cb.Items.AddRange(listNew.ToArray());
            //设置光标位置，否则光标位置始终保持在第一列，造成输入关键词的倒序排列
            cb.SelectionStart = cb.Text.Length;
            //保持鼠标指针原来状态，有时鼠标指针会被下拉框覆盖，所以要进行一次设置
            Cursor = Cursors.Default;
            //自动弹出下拉框
            cb.DroppedDown = true;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0112 && m.WParam.ToInt32() == 61490) return;
            base.WndProc(ref m);
        }

        //菜单栏子目录点击触发事件
        private void MenuChildClick(object sender, EventArgs e)
        {
            string formname = sender.ToString();
            if (formname.Equals("读取刀线"))
            {
                this.panel1.Controls.Clear();
                MyMIS.ReadCF2 report = new MyMIS.ReadCF2();
                report.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(report);
            }

            if (formname.Equals("重新登录"))
            {
                this.Hide();
                Application.EnableVisualStyles();
                Application.Exit();
                Application.Restart();
            }

            if (formname.Equals("退出"))
            {
                System.Environment.Exit(0);
            }

            if (formname.Equals("软件信息"))
            {
                this.panel1.Controls.Clear();
                CustomControl.Help help = new CustomControl.Help();
                help.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(help);
            }
            if (formname.Equals("操作文档"))
            {
                CustomControl.PDFHELP pdf = new CustomControl.PDFHELP();
                pdf.ShowDialog();
            }
            if (formname.Equals("权限管理"))
            {
                MenuBar.PermissionControl per = new MenuBar.PermissionControl();
                per.ShowDialog();
            }
            if (formname.Equals("书脊厚度计算"))
            {
                this.panel1.Controls.Clear();
                CustomControl.SpineCalculate sc = new CustomControl.SpineCalculate();
                sc.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(sc);
            }
            if (formname.Equals("纸价吨令转换"))
            {
                this.panel1.Controls.Clear();
                CustomControl.PaperPriceTonConversion pptc = new CustomControl.PaperPriceTonConversion();
                pptc.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(pptc);
            }
            if (formname.Equals("制版纸张规格尺寸表"))
            {
                this.panel1.Controls.Clear();
                CustomControl.StandardSizeTable ss = new CustomControl.StandardSizeTable();
                ss.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(ss);
            }
            if (formname.Equals("产品与原纸换算"))
            {
                this.panel1.Controls.Clear();
                CustomControl.ConversionProductOriginalPaper cpop = new CustomControl.ConversionProductOriginalPaper();
                cpop.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(cpop);
            }
            if (formname.Equals("基础单位换算"))
            {
                this.panel1.Controls.Clear();
                CustomControl.UnitConversion uc = new CustomControl.UnitConversion();
                uc.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(uc);
            }
            if (formname.Equals("薄膜吨令转换"))
            {
                this.panel1.Controls.Clear();
                CustomControl.ThinFilmTonerConversion ttc = new CustomControl.ThinFilmTonerConversion();
                ttc.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(ttc);
            }
            if (formname.Equals("印刷开版工具"))
            {
                this.panel1.Controls.Clear();
                CustomControl.PrintingTool pt = new CustomControl.PrintingTool();
                pt.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(pt);
            }
            if (formname.Equals("隐藏树目录"))
            {
                splitContainer1.Panel1Collapsed = true;
            }
            if (formname.Equals("显示树目录"))
            {
                splitContainer1.Panel1Collapsed = false;
            }
            if (formname.Equals("全展开树目录"))
            {
                treeView1.ExpandAll();
            }
            if (formname.Equals("全合并树目录"))
            {
                treeView1.CollapseAll();
            }
            if (formname.Equals("默认树目录"))
            {
                treeView1.CollapseAll();
                ExpandTree2();
            }
            if (formname.Equals("更新本地基础资料"))
            {
                this.Cursor = Cursors.WaitCursor;//等待
                try
                {
                    SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
                    string s = connect.GetConnectStr("ERP");
                    string connectionString = s;

                    #region 更新工序基础资料 
                    ////获取本地sqlite数据库信息，返回dt100
                    //DataTable dt100 = new DataTable();
                    //dt100 = Tools.Common.SQLiteHelper.ExecuteDatatable("select * from BasicProcessInformation");
                    ////获取time文本时间，根据时间查询SQL数据库，返回dt101
                    //DataTable dt101 = new DataTable();
                    //using (SqlConnection conn = new SqlConnection(connectionString))
                    //{
                    //    conn.Open();
                    //    using (SqlCommand cmd = conn.CreateCommand())
                    //    {
                    //        cmd.CommandType = CommandType.StoredProcedure;
                    //        cmd.CommandText = "ProcessOfConversion1";
                    //        SqlParameter[] para ={
                    //                   new SqlParameter("@times",SqlDbType.DateTime)};
                    //        para[0].Value = ReadDateTime();
                    //        try
                    //        {
                    //            cmd.CommandTimeout = 60 * 60 * 1000;
                    //            cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                    //            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    //            adapter.Fill(dt101);
                    //            SqlHelper.GetConnection().Close();
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            MessageBox.Show(ex.Message);
                    //            SqlHelper.GetConnection().Close();
                    //        }
                    //    }
                    //}
                    ////将两个datatable存入sqlite数据库成立两张临时表，dt1与dt2内连接取交际，返回数据表，更新本地数据库信息
                    //DataTable dt102 = new DataTable();
                    //DataColumn[] dc1 = new DataColumn[] { d10, d11, d12, d13, d14, d15 };
                    //DataColumn[] dc2 = new DataColumn[] { d20, d21, d22, d23, d24, d25 };
                    //dt102 = JoinTwoTable(dt100, dt101, dc2, dc1);
                    //DataTable dt103 = GenerateDataGX(dt102);
                    //string del2 = "delete from BasicProcessInformation";
                    //Tools.Common.SQLiteHelper.ExecuteNonQuery(Tools.Common.SQLiteHelper.CreateCommand(del2));
                    ////将更新后的数据还原到本地数据库
                    //InsertGXData(dt103,true);
                    #endregion

                    #region 更新物料基础资料
                    ////获取本地sqlite数据库信息，返回dt200
                    //DataTable dt200 = new DataTable();
                    //dt200 = Tools.Common.SQLiteHelper.ExecuteDatatable("select * from BasicMaterials");
                    ////获取time文本时间，根据时间查询sqlserver数据库，返回dt201
                    //DataTable dt201 = new DataTable();
                    //using (SqlConnection conn = new SqlConnection(connectionString))
                    //{
                    //    conn.Open();
                    //    using (SqlCommand cmd = conn.CreateCommand())
                    //    {
                    //        cmd.CommandType = CommandType.StoredProcedure;
                    //        cmd.CommandText = "MaterialQuery2";
                    //        SqlParameter[] para ={
                    //                   new SqlParameter("@times",SqlDbType.DateTime)};
                    //        para[0].Value = ReadDateTime();
                    //        try
                    //        {
                    //            cmd.CommandTimeout = 60 * 60 * 1000;
                    //            cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                    //            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    //            adapter.Fill(dt201);
                    //            SqlHelper.GetConnection().Close();
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            MessageBox.Show(ex.Message);
                    //            SqlHelper.GetConnection().Close();
                    //        }
                    //    }
                    //}
                    ////dt200与dt201在内存中内连接，返回datatable,并更新本地数据表
                    //DataTable dt202 = new DataTable();
                    //DataColumn[] dc3 = new DataColumn[] { d30, d31, d32, d33, d34, d35,d36 };
                    //DataColumn[] dc4 = new DataColumn[] { d40, d41, d42, d43, d44, d45, d46 };
                    //dt202 = JoinTwoTable(dt200, dt201, dc4, dc3);
                    //DataTable dt203 = GenerateDataWL(dt202);
                    //string del1 = "delete from BasicMaterials";
                    //Tools.Common.SQLiteHelper.ExecuteNonQuery(Tools.Common.SQLiteHelper.CreateCommand(del1));
                    ////将更新后的数据还原到本地数据库
                    //InsertWLData(dt203,true);
                    #endregion

                    #region 插入工序基础资料
                    InsertGXData(null,false);
                    #endregion

                    #region 插入物料基础资料
                    InsertWLData(null,false);
                    #endregion

                    WriteDateTime();
                    MessageBox.Show("本地数据库更新成功！","提示",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("本地数据库更新失败！","提示",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    MessageBox.Show(ex.Message);    
                }
                this.Cursor = Cursors.Default;//正常        
            }
            if (formname.Equals("换算工序正数"))
            {
                ERPInquire.MenuBar.ProcessConversion pc = ERPInquire.MenuBar.ProcessConversion.CreateInstrance();
                pc.Show();
            }
            if (formname.Equals("材料领用单据编辑"))
            {
                ERPInquire.MenuBar.MaterialsOfRecipients mor = ERPInquire.MenuBar.MaterialsOfRecipients.CreateInstrance();
                mor.Show();
            }
            if (formname.Equals("研发部产品档案批量更改物料"))
            {
                ERPInquire.MenuBar.ChangeMaterialY ci = ERPInquire.MenuBar.ChangeMaterialY.CreateInstrance();
                ci.Username = userName;
                ci.Userbumen = getuserdepartment();
                ci.Show();
            }
            if (formname.Equals("计划部生产施工单批量更改物料"))
            {
                ERPInquire.MenuBar.ChangeMaterialJ ci = ERPInquire.MenuBar.ChangeMaterialJ.CreateInstrance();
                ci.Username = userName;
                ci.Userbumen = getuserdepartment();
                ci.Show();
            }
            if (formname.Equals("财务税率修改"))
            {
                ERPInquire.MenuBar.Financialtax ft = ERPInquire.MenuBar.Financialtax.CreateInstrance();
                ft.Show();
            }
        }
        //树形子目录点击触发事件
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Level == 2)
            {
                comboBox1.Text = e.Node.Text;
                toolStripStatusLabel2.Text = "模块名称：" + e.Node.Text;
                if (userID.Equals("00000"))
                {
                }
                else
                {
                    this.panel1.Controls.Clear();
                    DataTable table = ds2.Tables["TABLE"];
                    string expression;
                    expression = "孙模块名称 ='" + e.Node.Text + "'";
                    DataRow[] foundRows;
                    //使用选择方法来找到匹配的所有行。
                    foundRows = table.Select(expression);
                    //过滤行,找到所要的行。
                    string strTemp1 = foundRows[0]["存储过程名称"].ToString();
                    string strTemp2 = foundRows[0]["模块类别"].ToString();
                    string strTemp3 = foundRows[0]["关键字"].ToString();
                    string strTemp4 = foundRows[0]["时间关键字"].ToString();
                    string strTemp5 = foundRows[0]["关键字1"].ToString();

                    if (strTemp1 != "" && strTemp2 != "")
                    {
                        #region 前后日期以及关键字查询

                        if (strTemp2.Equals("1"))
                        {
                            CustomControl.CommonControl2 cmd2 = new CustomControl.CommonControl2();
                            cmd2.ProName = strTemp1;
                            cmd2.TheKeyword = strTemp3;
                            cmd2.TimeKey = strTemp4;
                            cmd2.Dock = DockStyle.Fill;
                            cmd2.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd2);
                        }

                        #endregion 前后日期以及关键字查询

                        #region 关键字查询

                        if (strTemp2.Equals("2"))
                        {
                            CustomControl.CommonControl1 cmd1 = new CustomControl.CommonControl1();
                            cmd1.ProName = strTemp1;
                            cmd1.TheKeyword = strTemp3;
                            cmd1.Dock = DockStyle.Fill;
                            cmd1.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd1);
                        }

                        #endregion 关键字查询

                        #region 批量查询

                        if (strTemp2.Equals("3"))
                        {
                            try
                            {
                                SqlHelper.GetConnection();
                                DataTable dt = new DataTable();
                                string s1 = "select * from ERPInquire_M_BatchQueryData where ModuleNameChild='" + e.Node.Text + "'";
                                dt = SqlHelper.ExecuteDataTable(s1);
                                SqlHelper.GetConnection().Close();
                                if (dt.Rows.Count == 1)
                                {
                                    string nameTemp = dt.Rows[0][2].ToString();
                                    string arributeTemp = dt.Rows[0][3].ToString();
                                    string[] nameTemp1 = nameTemp.Split(',');
                                    string[] arributeTemp1 = arributeTemp.Split(',');
                                    CustomControl.CommonControl3 cmd3 = new CustomControl.CommonControl3();
                                    cmd3.names = nameTemp1;
                                    cmd3.attributes = arributeTemp1;
                                    cmd3.nodename = e.Node.Text;
                                    cmd3.ProName = strTemp1;
                                    cmd3.NodeName = e.Node.Text;
                                    cmd3.Dock = DockStyle.Fill;
                                    this.panel1.Controls.Add(cmd3);
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }

                        #endregion 批量查询

                        #region 区间日期查询

                        if (strTemp2.Equals("4"))
                        {
                            CustomControl.CommonControl4 cmd4 = new CustomControl.CommonControl4();
                            cmd4.ProName = strTemp1;
                            cmd4.Dock = DockStyle.Fill;
                            cmd4.TimeKey = strTemp4;
                            cmd4.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd4);
                        }

                        #endregion 区间日期查询

                        #region 按月份统计

                        if (strTemp2.Equals("5"))
                        {
                            CustomControl.CommonControl5 cmd5 = new CustomControl.CommonControl5();
                            cmd5.ProName = strTemp1;
                            cmd5.Dock = DockStyle.Fill;
                            cmd5.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd5);
                        }

                        #endregion 按月份统计

                        #region 时间区间以及两个查询关键字
                        if (strTemp2.Equals("8"))
                        {
                            CustomControl.CommonControl8 cmd8 = new CustomControl.CommonControl8();
                            cmd8.ProName = strTemp1;
                            cmd8.TheKeyword = strTemp3;
                            cmd8.TheKeyword1 = strTemp5;
                            cmd8.TimeKey = strTemp4;
                            cmd8.Dock = DockStyle.Fill;
                            cmd8.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd8);
                        }
                        #endregion

                        #region 两个关键字
                        if (strTemp2.Equals("9"))
                        {
                            CustomControl.CommonControl9 cmd9 = new CustomControl.CommonControl9();
                            cmd9.ProName = strTemp1;
                            cmd9.TheKeyword = strTemp3;
                            cmd9.TheKeyword1 = strTemp5;
                            cmd9.Dock = DockStyle.Fill;
                            cmd9.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd9);
                        }
                        #endregion
               
                        #region 印刷排产
                        if (strTemp2.Equals("10"))
                        {
                            CustomControl.CommonControl10 cmd10 = new CustomControl.CommonControl10();
                            cmd10.Dock = DockStyle.Fill;
                            cmd10.NodeName = e.Node.Text;
                            cmd10.ProName = strTemp1;
                            cmd10.TimeKey = strTemp4;
                            this.panel1.Controls.Add(cmd10);
                        }
                        #endregion

                        #region 单日期查询
                        if (strTemp2.Equals("12"))
                        {
                            //CustomControl.CommonControl12 cmd12 = new CustomControl.CommonControl12();
                            //cmd12.ProName = strTemp1;
                            //cmd12.TimeKey = strTemp4;
                            //cmd12.Dock = DockStyle.Fill;
                            //cmd12.NodeName = e.Node.Text;
                            //this.panel1.Controls.Add(cmd12);
                        }
                        #endregion

                        if (strTemp2.Equals("13"))
                        {
                            CustomControl.CommonControl13 cmd13 = new CustomControl.CommonControl13();
                            cmd13.Dock = DockStyle.Fill;
                            cmd13.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd13);
                        }
                        #region 材料实际库存查询
                        if (strTemp2.Equals("15"))
                        {
                            CustomControl.CommonControl15 cmd15 = new CustomControl.CommonControl15();
                            cmd15.Dock = DockStyle.Fill;
                            cmd15.NodeName = e.Node.Text;
                            cmd15.ProName = strTemp1;
                            this.panel1.Controls.Add(cmd15);
                        }
                        #endregion
                    }
                    if (strTemp1 == "" && strTemp2 != "")
                    {
                        if (strTemp2.Equals("6"))
                        {
                            CustomControl.CommonControl6 cmd6 = new CustomControl.CommonControl6();
                            cmd6.Dock = DockStyle.Fill;
                            cmd6.TimeKey = strTemp4;
                            cmd6.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd6);
                        }
                        if (strTemp2.Equals("7"))
                        {
                            CustomControl.CommonControl7 cmd7 = new CustomControl.CommonControl7();
                            cmd7.Dock = DockStyle.Fill;
                            cmd7.TimeKey = strTemp4;
                            cmd7.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd7);
                        }
                        if (strTemp2.Equals("11"))
                        {
                            CustomControl.CommonControl12 cmd12 = new CustomControl.CommonControl12();
                            cmd12.Dock = DockStyle.Fill;
                            cmd12.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(cmd12);
                        }
                        if (strTemp2.Equals("12"))
                        {
                            filesNewData.Query qe = new filesNewData.Query();
                            qe.Dock = DockStyle.Fill;
                            qe.NodeName = e.Node.Text;
                            this.panel1.Controls.Add(qe);
                        }
                    }
                 }
            }
        }

        //时间装载状态栏
        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime dt = System.DateTime.Now;
            toolStripStatusLabel3.Text = dt.ToString();
            
        }
        //树节点点击后方法
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeView1.SelectedNode.Parent != null)
            {
                if (IsselectNode != null && (IsselectNode.Parent != treeView1.SelectedNode.Parent.Parent && IsselectNode != treeView1.SelectedNode.Parent) && IsselectNode.Parent != null)
                {
                    if (IsselectNode.Parent.IsExpanded == true && IsselectNode.Parent != treeView1.Nodes[0] && IsselectNode.Parent != treeView1.SelectedNode.Parent)
                    {
                        IsselectNode = IsselectNode.Parent;
                    }
                    IsselectNode.Collapse();
                }
            }
            e.Node.Expand();
            if (treeView1.SelectedNode != treeView1.Nodes[0])
            {
                IsselectNode = treeView1.SelectedNode;
            }
            else if (IsselectNode != null)
            {
                IsselectNode.Collapse();
            }
        }
        #endregion

        #region 方法
        //根据姓名查找部门
        private string getuserdepartment()
        {
            try
            {
                Com.Hui.iMRP.Utils.SqlHelper.connectionStr = "packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
                DataTable dttr = Com.Hui.iMRP.Utils.SqlHelper.ExecuteDataTable("select * from JCZL_LogDepartment where name='" + userName + "'");
                return dttr.Rows[0][2].ToString();
            }
            catch
            {
                return "未知部门";
            }
        }
        private void LoadToolStrip2(DataSet ds)
        {
            string nameTemp = null;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string name = ds.Tables[0].Rows[i]["父模块名称"].ToString();
                if (nameTemp != name)
                {
                    ToolStripMenuItem Menu = new ToolStripMenuItem();
                    Menu.Text = name;
                    Menu.Font = new Font("微软雅黑", 10);
                    Menu.Image = new Bitmap("data/r5.ico");
                    this.menuStrip1.Items.Add(Menu);
                    DataRow[] row = ds.Tables[0].Select("父模块名称 = '" + name + "'");
                    for (int index = 0; index <= row.Length - 1; index++)
                    {
                        string MenuChild = Convert.ToString(row[index][1]);
                        ToolStripMenuItem i1 = new ToolStripMenuItem();
                        i1.Text = MenuChild;
                        i1.Font = new Font("微软雅黑",10);
                        i1.Image = new Bitmap("data/r4.ico");
                        i1.Click += new EventHandler(MenuChildClick);
                        Menu.DropDownItems.Add(i1);
                    }
                }
                nameTemp = name;
            }
        }
        private void LoadTree2(DataSet ds)
        {
            int ijk = 0;
            string nameTemp = null;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string name = ds.Tables[0].Rows[i]["树形模块名称"].ToString();
                if (nameTemp != name)
                {
                    DataRow[] row = ds.Tables[0].Select("树形模块名称 = '" + name + "'");
                    treeView1.ImageIndex = 0;
                    treeView1.Nodes.Add(name);
                    int mn = 0;
                    string nameTemp1 = null;
                    for (int index = 0; index <= row.Length - 1; index++)
                    {
                        TreeNode tn = treeView1.Nodes[ijk];
                        treeView1.Nodes[ijk].ImageIndex = 1;
                        string TreeChild = Convert.ToString(row[index][1]);
                        if (nameTemp1 != TreeChild)
                        {
                            tn.Nodes.Add(TreeChild);
                            DataRow[] row1 = ds.Tables[0].Select("树形模块名称 = '" + name + "' and 子模块名称='" + TreeChild + "'");
                            for (int index1 = 0; index1 <= row1.Length - 1; index1++)
                            {
                                TreeNode tn1 = treeView1.Nodes[ijk].Nodes[mn];
                                treeView1.Nodes[ijk].Nodes[mn].ImageIndex = 2;
                                string TreeChild1 = Convert.ToString(row1[index1][2]);
                                tn1.Nodes.Add(TreeChild1);
                            }
                            mn = mn + 1;
                        }
                        nameTemp1 = TreeChild;
                    }
                    ijk = ijk + 1;
                }
                nameTemp = name;
            }
            //展开树节点的所有目录
           //  treeView1.ExpandAll();
        }
        private void ExpandTree2()
        {
            for (int h = 0; h < treeView1.Nodes.Count; h++)
            {
                if (treeView1.Nodes[h].Level == 0)//判断当前节点的深度
                {
                    treeView1.Nodes[h].Expand();
                }
            }
        }   
        //读文本记录时间
        private string ReadDateTime()
        {
            if (!File.Exists("data/time.txt"))
            {
                return "1900-01-01 00:00:00";
            }
            else
            {
                FileStream fs = new FileStream("data/time.txt", FileMode.Open, FileAccess.Read);
                StreamReader sd = new StreamReader(fs);
                string str = sd.ReadLine();
                fs.Close();
                return str;
            }
        }
        //写入当前时间
        private void WriteDateTime()
        {
            FileStream fs = new FileStream("data/time.txt", FileMode.Create, FileAccess.Write);
            StreamWriter sr = new StreamWriter(fs);
            sr.WriteLine(DateTime.Now.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"));
            sr.Close();
            fs.Close();
        }
        /// <summary>
        /// 连接两个表
        /// </summary>
        /// <param name="First"></param>
        /// <param name="Second"></param>
        /// <param name="FJC"></param>
        /// <param name="SJC"></param>
        /// <returns></returns>
        public static DataTable JoinTwoTable(DataTable First, DataTable Second, string FJC, string SJC)
        {
            return JoinTwoTable(First, Second, new DataColumn[] { First.Columns[FJC] }, new DataColumn[] { First.Columns[SJC] });
        }
        /// <summary>
        /// 连接两个表
        /// </summary>
        /// <param name="First"></param>
        /// <param name="Second"></param>
        /// <param name="FJC"></param>
        /// <param name="SJC"></param>
        /// <returns></returns>
        protected static DataTable JoinTwoTable(DataTable First, DataTable Second, DataColumn FJC, DataColumn SJC)
        {
            return JoinTwoTable(First, Second, new DataColumn[] { FJC }, new DataColumn[] { SJC });
        }
        /// <summary>
        /// 连接两个Table
        /// </summary>
        /// <param name="First"></param>
        /// <param name="Second"></param>
        /// <param name="FJC"></param>
        /// <param name="SJC"></param>
        /// <returns></returns>
        protected static DataTable JoinTwoTable(DataTable First, DataTable Second, DataColumn[] FJC, DataColumn[] SJC)
        {
            //创建一个新的DataTable
            DataTable table = new DataTable("Join");
            using (DataSet ds = new DataSet())
            {
                //把DataTable Copy到DataSet中
                ds.Tables.AddRange(new DataTable[] { First.Copy(), Second.Copy() });
                DataColumn[] parentcolumns = new DataColumn[FJC.Length];
                for (int i = 0; i < parentcolumns.Length; i++)
                {
                    parentcolumns[i] = ds.Tables[0].Columns[FJC[i].ColumnName];
                }
                DataColumn[] childcolumns = new DataColumn[SJC.Length];
                for (int i = 0; i < childcolumns.Length; i++)
                {
                    childcolumns[i] = ds.Tables[1].Columns[SJC[i].ColumnName];
                }
                //创建关联
                DataRelation r = new DataRelation(string.Empty, parentcolumns, childcolumns, false);
                ds.Relations.Add(r);
                //为关联表创建列
                for (int i = 0; i < First.Columns.Count; i++)
                {
                    table.Columns.Add(First.Columns[i].ColumnName, First.Columns[i].DataType);
                }
                for (int i = 0; i < Second.Columns.Count; i++)
                {
                    //看看有没有重复的列，如果有在第二个DataTable的Column的列明后加_Second
                    if (!table.Columns.Contains(Second.Columns[i].ColumnName))
                        table.Columns.Add(Second.Columns[i].ColumnName, Second.Columns[i].DataType);
                    else
                        table.Columns.Add(Second.Columns[i].ColumnName + "_Second", Second.Columns[i].DataType);
                }
                table.BeginLoadData();
                foreach (DataRow firstrow in ds.Tables[0].Rows)
                {
                    //得到行的数据
                    DataRow[] childrows = firstrow.GetChildRows(r);
                    if (childrows != null && childrows.Length > 0)
                    {
                        object[] parentarray = firstrow.ItemArray;
                        foreach (DataRow secondrow in childrows)
                        {
                            object[] secondarray = secondrow.ItemArray;
                            object[] joinarray = new object[parentarray.Length + secondarray.Length];
                            Array.Copy(parentarray, 0, joinarray, 0, parentarray.Length);
                            Array.Copy(secondarray, 0, joinarray, parentarray.Length, secondarray.Length);
                            table.LoadDataRow(joinarray, true);
                        }
                    }
                }
                table.EndLoadData();
            }
           return table;
        }
        //将更新后的工序数据取出来
        private DataTable GenerateDataGX(DataTable dtBefore)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("产品编号", typeof(string)));
            dt.Columns.Add(new DataColumn("部件编号", typeof(string)));
            dt.Columns.Add(new DataColumn("部件名称", typeof(string)));
            dt.Columns.Add(new DataColumn("上机长", typeof(string)));
            dt.Columns.Add(new DataColumn("上机宽", typeof(string)));
            dt.Columns.Add(new DataColumn("正面颜色", typeof(string)));
            dt.Columns.Add(new DataColumn("反面颜色", typeof(string)));
            dt.Columns.Add(new DataColumn("印刷方式", typeof(string)));
            dt.Columns.Add(new DataColumn("得率", typeof(string)));
            dt.Columns.Add(new DataColumn("工序编码", typeof(string)));
            dt.Columns.Add(new DataColumn("工序类别", typeof(string)));
            dt.Columns.Add(new DataColumn("工序名称", typeof(string)));
            dt.Columns.Add(new DataColumn("变化系数", typeof(string)));
            //将传入的原始的表dtBefore中每一行中的每个数据复制给dr2这新行，再加入到新表dt中  
            DataRow dr2 = null;
            foreach (DataRow row in dtBefore.Rows)
            {
                dr2 = dt.NewRow();
                dr2[0] = row["产品编号"];
                dr2[1] = row["部件编号"];
                dr2[2] = row["部件名称"];
                dr2[3] = row["上机长"];
                dr2[4] = row["上机宽"];
                dr2[5] = row["正面颜色"];
                dr2[6] = row["反面颜色"];
                dr2[7] = row["印刷方式"];
                dr2[8] = row["得率"];
                dr2[9] = row["工序编码"];
                dr2[10] = row["工序类别"];
                dr2[11] = row["工序名称"];
                dr2[12] = row["变化系数"];
                dt.Rows.Add(dr2);
            }
            return dt;
        }
        //将更新后的物料数据取出来
        private DataTable GenerateDataWL(DataTable dtBefore)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("产品编号", typeof(string)));
            dt.Columns.Add(new DataColumn("产品名称", typeof(string)));
            dt.Columns.Add(new DataColumn("部件名称", typeof(string)));
            dt.Columns.Add(new DataColumn("物料大类", typeof(string)));
            dt.Columns.Add(new DataColumn("物料小类", typeof(string)));
            dt.Columns.Add(new DataColumn("物料编号", typeof(string)));
            dt.Columns.Add(new DataColumn("物料名称", typeof(string)));
            dt.Columns.Add(new DataColumn("上机长", typeof(string)));
            dt.Columns.Add(new DataColumn("上机宽", typeof(string)));
            dt.Columns.Add(new DataColumn("得率", typeof(string)));
            dt.Columns.Add(new DataColumn("标准用量", typeof(string)));
            //将传入的原始的表dtBefore中每一行中的每个数据复制给dr2这新行，再加入到新表dt中  
            DataRow dr2 = null;
            foreach (DataRow row in dtBefore.Rows)
            {
                dr2 = dt.NewRow();
                dr2[0] = row["产品编号"];
                dr2[1] = row["产品名称"];
                dr2[2] = row["部件名称"];
                dr2[3] = row["物料大类"];
                dr2[4] = row["物料小类"];
                dr2[5] = row["物料编号"];
                dr2[6] = row["物料名称"];
                dr2[7] = row["上机长"];
                dr2[8] = row["上机宽"];
                dr2[9] = row["得率"];
                dr2[10] = row["标准用量"];
                dt.Rows.Add(dr2);
            }
            return dt;
        }
        //插入工序基本信息表
        private void InsertGXData(DataTable dt,bool b)
        {
            SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
            string s = connect.GetConnectStr("ERP");
            string connectionString = s;

            if (b==true)
            {
                //根据比较结果更新本地数据库
                string connStr = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                using (SQLiteConnection con = new SQLiteConnection(connStr))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "INSERT INTO BasicProcessInformation (CPBH,ZBXH_BJBH,BJMC,SJC,SJK,ZMYS,FMYS,FSMC,DL,GXBH,GBLB,GXMC,BHCS) VALUES(@CPBH,@ZBXH_BJBH,@BJMC,@SJC,@SJK,@ZMYS,@FMYS,@FSMC,@DL,@GXBH,@GBLB,@GXMC,@BHCS)";
                    for (int n = 0; n < dt.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@ZBXH_BJBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@ZMYS", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@FMYS", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@FSMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GBLB", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BHCS", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@ZBXH_BJBH"].Value = dt.Rows[n]["部件编号"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt.Rows[n]["部件名称"].ToString();
                        cmd.Parameters["@SJC"].Value = dt.Rows[n]["上机长"].ToString();
                        cmd.Parameters["@SJK"].Value = dt.Rows[n]["上机宽"].ToString();
                        cmd.Parameters["@ZMYS"].Value = dt.Rows[n]["正面颜色"].ToString();
                        cmd.Parameters["@FMYS"].Value = dt.Rows[n]["反面颜色"].ToString();
                        cmd.Parameters["@FSMC"].Value = dt.Rows[n]["印刷方式"].ToString();
                        cmd.Parameters["@DL"].Value = dt.Rows[n]["得率"].ToString();
                        cmd.Parameters["@GXBH"].Value = dt.Rows[n]["工序编码"].ToString();
                        cmd.Parameters["@GBLB"].Value = dt.Rows[n]["工序类别"].ToString();
                        cmd.Parameters["@GXMC"].Value = dt.Rows[n]["工序名称"].ToString();
                        cmd.Parameters["@BHCS"].Value = dt.Rows[n]["变化系数"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
            }
            else 
            {
              //查询sql server数据库更新本地数据库       
              DataTable dt1 = new DataTable();
              using (SqlConnection conn = new SqlConnection(connectionString))
              {
                  conn.Open();
                  using (SqlCommand cmd = conn.CreateCommand())
                  {
                      cmd.CommandType = CommandType.StoredProcedure;
                      cmd.CommandText = "ProcessOfConversion";
                      SqlParameter[] para ={
                                       new SqlParameter("@times",SqlDbType.DateTime)};
                      para[0].Value = ReadDateTime();
                      try
                      {
                          cmd.CommandTimeout = 60 * 60 * 1000;
                          cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                          SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                          adapter.Fill(dt1);
                          SqlHelper.GetConnection().Close();
                      }
                      catch (Exception ex)
                      {
                          MessageBox.Show(ex.Message);
                          SqlHelper.GetConnection().Close();
                      }
                  }
              }
                //删除数据
                string connStr = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                using (SQLiteConnection con = new SQLiteConnection(connStr))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "delete from  BasicProcessInformation where CPBH=@CPBH and ZBXH_BJBH=@ZBXH_BJBH and BJMC=@BJMC and GXBH=@GXBH and GBLB=@GBLB and GXMC=@GXMC";
                    for (int n = 0; n < dt1.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@ZBXH_BJBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GBLB", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXMC", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt1.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@ZBXH_BJBH"].Value = dt1.Rows[n]["部件编号"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt1.Rows[n]["部件名称"].ToString();
                        cmd.Parameters["@GXBH"].Value = dt1.Rows[n]["工序编码"].ToString();
                        cmd.Parameters["@GBLB"].Value = dt1.Rows[n]["工序类别"].ToString();
                        cmd.Parameters["@GXMC"].Value = dt1.Rows[n]["工序名称"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
              //新增数据
              using (SQLiteConnection con = new SQLiteConnection(connStr))
              {
                  con.Open();
                  DbTransaction trans = con.BeginTransaction();//开始事务       
                  SQLiteCommand cmd = new SQLiteCommand(con);
                  cmd.CommandText = "INSERT INTO BasicProcessInformation (CPBH,ZBXH_BJBH,BJMC,SJC,SJK,ZMYS,FMYS,FSMC,DL,GXBH,GBLB,GXMC,BHCS) VALUES(@CPBH,@ZBXH_BJBH,@BJMC,@SJC,@SJK,@ZMYS,@FMYS,@FSMC,@DL,@GXBH,@GBLB,@GXMC,@BHCS)";
                  for (int n = 0; n < dt1.Rows.Count; n++)
                  {
                      cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@ZBXH_BJBH", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@ZMYS", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@FMYS", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@FSMC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@GXBH", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@GBLB", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@GXMC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@BHCS", DbType.String));
                      cmd.Parameters["@CPBH"].Value = dt1.Rows[n]["产品编号"].ToString();
                      cmd.Parameters["@ZBXH_BJBH"].Value = dt1.Rows[n]["部件编号"].ToString();
                      cmd.Parameters["@BJMC"].Value = dt1.Rows[n]["部件名称"].ToString();
                      cmd.Parameters["@SJC"].Value = dt1.Rows[n]["上机长"].ToString();
                      cmd.Parameters["@SJK"].Value = dt1.Rows[n]["上机宽"].ToString();
                      cmd.Parameters["@ZMYS"].Value = dt1.Rows[n]["正面颜色"].ToString();
                      cmd.Parameters["@FMYS"].Value = dt1.Rows[n]["反面颜色"].ToString();
                      cmd.Parameters["@FSMC"].Value = dt1.Rows[n]["印刷方式"].ToString();
                      cmd.Parameters["@DL"].Value = dt1.Rows[n]["得率"].ToString();
                      cmd.Parameters["@GXBH"].Value = dt1.Rows[n]["工序编码"].ToString();
                      cmd.Parameters["@GBLB"].Value = dt1.Rows[n]["工序类别"].ToString();
                      cmd.Parameters["@GXMC"].Value = dt1.Rows[n]["工序名称"].ToString();
                      cmd.Parameters["@BHCS"].Value = dt1.Rows[n]["变化系数"].ToString();
                      cmd.ExecuteNonQuery();
                  }
                  trans.Commit();//提交事务  
                  con.Close();
              }
            }
        }
        //插入物料基本信息表
        private void InsertWLData(DataTable dt, bool b)
        {
            SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
            string s = connect.GetConnectStr("ERP");
            string connectionString = s;

            if (b==true)
            {
                //根据比较结果更新本地数据库
                string connStr1 = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                using (SQLiteConnection con = new SQLiteConnection(connStr1))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "INSERT INTO BasicMaterials (CPBH,CPMC,BJMC,SJC,SJK,DL,WLDL,WLZL,WLBH,WLMC,BZYL) VALUES(@CPBH,@CPMC,@BJMC,@SJC,@SJK,@DL,@WLDL,@WLZL,@WLBH,@WLMC,@BZYL)";
                    for (int n = 0; n < dt.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLDL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLZL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BZYL", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@CPMC"].Value = dt.Rows[n]["产品名称"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt.Rows[n]["部件名称"].ToString();
                        cmd.Parameters["@SJC"].Value = dt.Rows[n]["上机长"].ToString();
                        cmd.Parameters["@SJK"].Value = dt.Rows[n]["上机宽"].ToString();
                        cmd.Parameters["@DL"].Value = dt.Rows[n]["得率"].ToString();
                        cmd.Parameters["@WLDL"].Value = dt.Rows[n]["物料大类"].ToString();
                        cmd.Parameters["@WLZL"].Value = dt.Rows[n]["物料小类"].ToString();
                        cmd.Parameters["@WLBH"].Value = dt.Rows[n]["物料编号"].ToString();
                        cmd.Parameters["@WLMC"].Value = dt.Rows[n]["物料名称"].ToString();
                        cmd.Parameters["@BZYL"].Value = dt.Rows[n]["标准用量"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
            }
            else
            {
              //查询sql server数据库更新本地数据库       
              DataTable dt2 = new DataTable();
              using (SqlConnection conn = new SqlConnection(connectionString))
              {
                  conn.Open();
                  using (SqlCommand cmd = conn.CreateCommand())
                  {
                      cmd.CommandType = CommandType.StoredProcedure;
                      cmd.CommandText = "MaterialQuery1";
                      SqlParameter[] para = { new SqlParameter("@times", SqlDbType.DateTime) };
                      para[0].Value = ReadDateTime();
                      try
                      {
                          cmd.CommandTimeout = 60 * 60 * 1000;
                          cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                          SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                          adapter.Fill(dt2);
                          SqlHelper.GetConnection().Close();
                      }
                      catch (Exception ex)
                      {
                          MessageBox.Show(ex.Message);
                          SqlHelper.GetConnection().Close();
                      }
                  }
              }
              string connStr1 = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                //删除数据
                using (SQLiteConnection con = new SQLiteConnection(connStr1))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "delete from  BasicMaterials where CPBH=@CPBH and CPMC=@CPMC and BJMC=@BJMC and WLDL=@WLDL and WLZL=@WLZL and WLBH=@WLBH and WLMC=@WLMC";
                    for (int n = 0; n < dt2.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLDL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLZL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLMC", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt2.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@CPMC"].Value = dt2.Rows[n]["产品名称"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt2.Rows[n]["部件名称"].ToString();
                        cmd.Parameters["@WLDL"].Value = dt2.Rows[n]["物料大类"].ToString();
                        cmd.Parameters["@WLZL"].Value = dt2.Rows[n]["物料小类"].ToString();
                        cmd.Parameters["@WLBH"].Value = dt2.Rows[n]["物料编号"].ToString();
                        cmd.Parameters["@WLMC"].Value = dt2.Rows[n]["物料名称"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
                //插入数据
                using (SQLiteConnection con = new SQLiteConnection(connStr1))
              {
                  con.Open();
                  DbTransaction trans = con.BeginTransaction();//开始事务       
                  SQLiteCommand cmd = new SQLiteCommand(con);
                  cmd.CommandText = "INSERT INTO BasicMaterials (CPBH,CPMC,BJMC,SJC,SJK,DL,WLDL,WLZL,WLBH,WLMC,BZYL) VALUES(@CPBH,@CPMC,@BJMC,@SJC,@SJK,@DL,@WLDL,@WLZL,@WLBH,@WLMC,@BZYL)";
                  for (int n = 0; n < dt2.Rows.Count; n++)
                  {
                      cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@WLDL", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@WLZL", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@WLBH", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@WLMC", DbType.String));
                      cmd.Parameters.Add(new SQLiteParameter("@BZYL", DbType.String));
                      cmd.Parameters["@CPBH"].Value = dt2.Rows[n]["产品编号"].ToString();
                      cmd.Parameters["@CPMC"].Value = dt2.Rows[n]["产品名称"].ToString();
                      cmd.Parameters["@BJMC"].Value = dt2.Rows[n]["部件名称"].ToString();
                      cmd.Parameters["@SJC"].Value = dt2.Rows[n]["上机长"].ToString();
                      cmd.Parameters["@SJK"].Value = dt2.Rows[n]["上机宽"].ToString();
                      cmd.Parameters["@DL"].Value = dt2.Rows[n]["得率"].ToString();
                      cmd.Parameters["@WLDL"].Value = dt2.Rows[n]["物料大类"].ToString();
                      cmd.Parameters["@WLZL"].Value = dt2.Rows[n]["物料小类"].ToString();
                      cmd.Parameters["@WLBH"].Value = dt2.Rows[n]["物料编号"].ToString();
                      cmd.Parameters["@WLMC"].Value = dt2.Rows[n]["物料名称"].ToString();
                      cmd.Parameters["@BZYL"].Value = dt2.Rows[n]["标准用量"].ToString();
                      cmd.ExecuteNonQuery();
                  }
                  trans.Commit();//提交事务  
                  con.Close();
              }
            }      
        }
        #endregion 方法
    }
}