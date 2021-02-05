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
    public partial class PersonnelAccess : Form
    {
        #region 变量
        public string name1;
        public string name2;
        public DataSet ds;

        bool beginMove = false;
        int currentXPosition;
        int currentYPosition;
        #endregion

        #region 定量
        private static PersonnelAccess frm = null;
        #endregion

        #region 初始化
        private PersonnelAccess()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
            init();
        }

        public static PersonnelAccess CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new PersonnelAccess();
            }
            return frm;
        }
        private void init()
        {
            string sql = "select PermissionsName AS 权限组名字 from ERPInquire_S_StaffTableM ";
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string temp = Convert.ToString(dt.Rows[i]["权限组名字"].ToString());
                checkedListBox1.Items.Add(temp);
            }  
        }
        private void PersonnelAccess_Load(object sender, EventArgs e)
        {
            label1.Text = name1;
            this.TopMost = true;
            FillCheckBoxList(name2, checkedListBox1);
        }
        public void FillCheckBoxList(string str, CheckedListBox checkBoxList)
        {
            for (int j = 0; j < checkedListBox1.Items.Count; j++)
            {
                if(checkedListBox1.Items[j].ToString()==str)
                  checkedListBox1.SetItemChecked(j, true);
            }
        }
        #endregion

        #region 窗体事件
        //保存
        private void button1_Click_1(object sender, EventArgs e)
        {
            //string sql=
            // string sql1 = checkedListBox1.SelectedItem.ToString();
            // SqlHelper.ExecCommand();
            string output = string.Empty;
            for (int i = 0; i < checkedListBox1.CheckedIndices.Count; i++)
            {
                output += checkedListBox1.Items[
                checkedListBox1.CheckedIndices[i]].ToString();
            }
            string sql2 = "SELECT PermissionsID as 权限号,PermissionsName AS 权限组名称 from ERPInquire_S_StaffTableM";
            DataTable dttemp = new DataTable();
            dttemp=SqlHelper.ExecuteDataTable(sql2);
            SqlHelper.GetConnection().Close();
            string expression;
            expression = "权限组名称 ='" + output + "'";
            DataRow[] foundRows;
            //使用选择方法来找到匹配的所有行。
            foundRows = dttemp.Select(expression);
            if (foundRows.Length == 1)
            {
                //过滤行,找到所要的行。
                string strTemp1 = foundRows[0]["权限号"].ToString();
                string sql3 = "update  ZCX_YHQX_M set ERP_Authen ='" + strTemp1 + "' where name='" + name1 + "'";
                SqlHelper.ExecCommand(sql3);
                SqlHelper.GetConnection().Close();
            }
            else
            {
                string sql3 = "update  ZCX_YHQX_M set ERP_Authen ='' where name='" + name1 + "'";
                SqlHelper.ExecCommand(sql3);
                SqlHelper.GetConnection().Close();
            }
            this.Close();
        }
        //返回
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.CurrentValue == CheckState.Checked) return;//取消选中就不用进行以下操作  
            for (int i = 0; i < ((CheckedListBox)sender).Items.Count; i++)
            {
              ((CheckedListBox)sender).SetItemChecked(i, false);//将所有选项设为不选中  
            }
            e.NewValue = CheckState.Checked;//刷新  
        }
        #endregion

        #region  方法

        #endregion

        #region 鼠标事件
        private void PersonnelAccess_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                beginMove = true;
                currentXPosition = MousePosition.X;//鼠标的x坐标为当前窗体左上角x坐标
                currentYPosition = MousePosition.Y;//鼠标的y坐标为当前窗体左上角y坐标
            }
        }

        private void PersonnelAccess_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                currentXPosition = 0; //设置初始状态
                currentYPosition = 0;
                beginMove = false;
            }
        }
        
        private void PersonnelAccess_MouseMove(object sender, MouseEventArgs e)
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
