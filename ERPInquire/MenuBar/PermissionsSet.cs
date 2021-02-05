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
    public partial class PermissionsSet : Form
    {
        #region 定量
        public string name1;
        bool beginMove = false;
        int currentXPosition;
        int currentYPosition;
        #endregion

        #region 变量
        private static PermissionsSet frm = null;
        #endregion

        #region 初始化
        private PermissionsSet()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }

        public static PermissionsSet CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new PermissionsSet();
            }
            return frm;
        }
        private void init()
        {
            string sql = "select PermissionControl2 AS 权限字符串 from ERPInquire_S_StaffTableM where PermissionsName='"+name1+"'";
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            string temp = Convert.ToString(dt.Rows[0]["权限字符串"].ToString());
            SqlHelper.GetConnection();
            DataSet ds = new DataSet();
            string[] condititons = temp.Split(',');
            string ss = "SELECT ModuleNameChild as 名称 FROM ERPInquire_S_TreeMenuBarControlD1_D1";
            ds = SqlHelper.ExecuteDataSet(ss, null);
            SqlHelper.GetConnection().Close();
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string temp1 = Convert.ToString(ds.Tables[0].Rows[i]["名称"].ToString());
                checkedListBox1.Items.Add(temp1);
            }
            FillCheckBoxList(temp,checkedListBox1);
        }
        #endregion

        #region 窗体事件
        private void PermissionsSet_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            init();
            label1.Text = name1;
        }
        public void FillCheckBoxList(string str, CheckedListBox checkBoxList)
        {
            DataSet ds = new DataSet();
            string[] condititons = str.Split(',');
            string ss = "SELECT ModuleNameChild as 名称 FROM ERPInquire_S_TreeMenuBarControlD1_D1  WHERE  PermissionControl IN (";
            for (int i = 0; i < condititons.Length; i++)
            {
                ss = ss + condititons[i] + ",";
            }
            ss = ss.Substring(0, ss.Length - 1);
            ss = ss + ") ";
            ds = SqlHelper.ExecuteDataSet(ss, null);
            SqlHelper.GetConnection().Close();
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string temp = Convert.ToString(ds.Tables[0].Rows[i]["名称"].ToString());
                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                {
                    if (checkedListBox1.Items[j].ToString() == temp)
                        checkedListBox1.SetItemChecked(j, true);
                }
            }
        }
        //取消
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        //保存
        private void button1_Click(object sender, EventArgs e)
        {
           string str = string.Empty;
           for (int i = 0; i < checkedListBox1.Items.Count; i++)
           {
               if (checkedListBox1.GetItemChecked(i))
               {
                    str = str+checkedListBox1.GetItemText(checkedListBox1.Items[i]) + ",";
               }
            }
            str = str.Substring(0, str.Length - 1);
            DataSet ds = new DataSet();
            string[] condititons = str.Split(',');
            string ss = "SELECT PermissionControl as 权限 FROM ERPInquire_S_TreeMenuBarControlD1_D1 WHERE ModuleNameChild IN (";
            for (int i = 0; i < condititons.Length; i++)
            {
                ss = ss + "'"+condititons[i]+"'"+ ",";
            }
            ss = ss.Substring(0, ss.Length - 1);
            ss = ss + ") ";
            ds = SqlHelper.ExecuteDataSet(ss, null);
            SqlHelper.GetConnection().Close();
            string[] s23 = new string[ds.Tables[0].Rows.Count];
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                try
                {
                    s23[i] = ds.Tables[0].Rows[i]["权限"].ToString();
                }
                catch (Exception)
                {
                    continue;
                }
            }
            string str332=Toheavy(s23);
            string sql332 = "update ERPInquire_S_StaffTableM set PermissionControl2='" +str332+ "' where PermissionsName='"+name1+"'";
            SqlHelper.ExecCommand(sql332);
            SqlHelper.GetConnection().Close();
            this.Close();
        }
        #endregion

        #region 方法
        //字符串去重
          private string Toheavy(string[] st)
        {
           string str = string.Empty;
           List<string> listString = new List<string>();
           foreach (string eachString in st) 
           {
               if (!listString.Contains(eachString))
               listString.Add(eachString);
           }
           foreach (string eachString in listString)
           {
                str = str + eachString + ",";
           }
            str = str.Substring(0, str.Length - 1);
            return str;
        }
       #endregion

       #region 鼠标事件
       private void PermissionsSet_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                beginMove = true;
                currentXPosition = MousePosition.X;//鼠标的x坐标为当前窗体左上角x坐标
                currentYPosition = MousePosition.Y;//鼠标的y坐标为当前窗体左上角y坐标
            }
        }

        private void PermissionsSet_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                currentXPosition = 0; //设置初始状态
                currentYPosition = 0;
                beginMove = false;
            }
        }

        private void PermissionsSet_MouseMove(object sender, MouseEventArgs e)
        {
            if (beginMove)
            {
                this.Left += MousePosition.X - currentXPosition;//根据鼠标x坐标确定窗体的左边坐标x
                this.Top += MousePosition.Y - currentYPosition;//根据鼠标的y坐标窗体的顶部，即Y坐标
                currentXPosition = MousePosition.X;
                currentYPosition = MousePosition.Y;
            }
        }
        #endregion
    }
}
