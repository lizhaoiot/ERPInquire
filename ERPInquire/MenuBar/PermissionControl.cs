using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERPInquire.MenuBar
{
    public partial class PermissionControl : Form
    {
        #region 定量
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        #endregion

        #region 变量

        #endregion

        #region 初始化
        public PermissionControl()
        {
            InitializeComponent();
        }
        private void PermissionControl_Load(object sender, EventArgs e)
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            DisplayItem();
            DisplayList2();
            DisplayList3();
        }
        #endregion

        #region 用户组权限查询
        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            int index = listBox1.IndexFromPoint(e.X, e.Y);
            listBox1.SelectedIndex = index;
            if (listBox1.SelectedIndex != -1)
            {
                try
                {
                    string expression;
                    expression = "姓名 ='" + listBox1.SelectedItem.ToString() + "'";
                    DataRow[] foundRows;
                    //使用选择方法来找到匹配的所有行。
                    foundRows = dt.Select(expression);
                    //过滤行,找到所要的行。
                    string strTemp1 = foundRows[0]["权限组"].ToString();
                    string strTemp2 = foundRows[0]["权限组名字"].ToString();
                    SqlHelper.GetConnection();
                    DataSet ds = new DataSet();
                    string[] condititons = strTemp1.Split(',');
                    string ss = "SELECT ModuleNameChild as 模块名字 from ERPInquire_S_TreeMenuBarControlD1_D1 WHERE PermissionControl IN (";
                    for (int i = 0; i < condititons.Length; i++)
                    {
                        ss = ss + condititons[i] + ",";
                    }
                    ss = ss.Substring(0, ss.Length - 1);
                    ss = ss + ")  ";
                    ds = SqlHelper.ExecuteDataSet(ss, null);
                    SqlHelper.GetConnection().Close();
                    Display dp = Display.CreateInstrance();
                    dp.ds = ds;
                    dp.s1 = strTemp2;
                    dp.Show();
                }
                catch
                {
                    MessageBox.Show("此用户未分配权限组");
                }
            }
        }

        private void DisplayItem()
        {
            string sql = "select M.name as 姓名,M.ERP_Authen as 权限管理,D1.PermissionControl2 AS 权限组,D1.PermissionsName AS 权限组名字 from ZCX_YHQX_M M left join ERPInquire_S_StaffTableM D1 on M.ERP_Authen=D1.PermissionsID";
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            for (int i = 0; i < dt.Rows.Count; i++)
                this.listBox1.Items.Add(Convert.ToString(dt.Rows[i]["姓名"].ToString()));
        }

        #endregion

        #region 人员权限管理
        private void DisplayList2()
        {
            listView2.GridLines = true;//表格是否显示网格线
            listView2.FullRowSelect = true;//是否选中整行
            listView2.View = View.Details;//设置显示方式
            listView2.Scrollable = true;//是否自动显示滚动条
            string sql = "select m.name as 姓名,D1.PermissionsName AS 权限组名称 from ZCX_YHQX_M M left join ERPInquire_S_StaffTableM D1 on m.ERP_Authen=D1.PermissionsID";
            dt1 = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                ListViewItem[] p = new ListViewItem[1];
                p[0] = new ListViewItem(new string[] { Convert.ToString(dt1.Rows[i]["姓名"].ToString()), Convert.ToString(dt1.Rows[i]["权限组名称"].ToString()) });
                this.listView2.Items.AddRange(p);
            }
        }
        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
            ListView.SelectedIndexCollection indexes = this.listView2.SelectedIndices;
            if (indexes.Count > 0)
            {
                int index = indexes[0];
                string sPartNo = this.listView2.Items[index].SubItems[0].Text;//获取第一列的值  
                string sPartName = this.listView2.Items[index].SubItems[1].Text;//获取第二列的值 
                PersonnelAccess pa = PersonnelAccess.CreateInstrance();
                pa.name1 = sPartNo;
                pa.name2 = sPartName;
                pa.Show();
            }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("操作失败！\n" + ex.Message, "提示", MessageBoxButtons.OK,
            //        MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            //}
        }
        //刷新
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            listView2.Items.Clear();
            DisplayList2();
        }
        //保存
        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }
        #endregion

        #region 权限组所属模块管理
        private void DisplayList3()
        {
            listView3.GridLines = true;//表格是否显示网格线
            listView3.FullRowSelect = true;//是否选中整行
            listView3.View = View.Details;//设置显示方式
            listView3.Scrollable = true;//是否自动显示滚动条
            string sql = "select PermissionsName AS 权限组名称 from ERPInquire_S_StaffTableM";
            dt1 = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                ListViewItem[] p = new ListViewItem[1];
                p[0] = new ListViewItem(new string[] { Convert.ToString(dt1.Rows[i]["权限组名称"].ToString())});
                this.listView3.Items.AddRange(p);
            }
        }
        private void listView3_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection indexes = this.listView3.SelectedIndices;
            if (indexes.Count > 0)
            {
                int index = indexes[0];
                string sPartNo = this.listView3.Items[index].SubItems[0].Text;//获取第一列的值  
                PermissionsSet ps = PermissionsSet.CreateInstrance();
                ps.name1 = sPartNo;
                ps.Show();
            }
        }
        #endregion


    }
}
