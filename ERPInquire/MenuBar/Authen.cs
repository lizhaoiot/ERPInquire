using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERPInquire.MenuBar
{
    public partial class Authen : Form
    {
        #region 定量

        #endregion

        #region 变量

        SqlConnect.UserList user;

        #endregion

        #region 初始化

        public Authen()
        {
            InitializeComponent();
            this.ControlBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        }

        #endregion

        #region 窗体事件
        private void Authen_Load(object sender, EventArgs e)
        {
            load();
        }
        //确定按钮
        private void button1_Click(object sender, EventArgs e)
        {
            saveauth();
            this.Hide();
            Login l = new Login();
            l.ShowDialog();
        }
        //刷新XML菜单
        private void button2_Click(object sender, EventArgs e)
        {
            //删除本地XML文件
            DeleteFiles();
        }
        #endregion

        #region 方法
        private void DeleteFiles()
        {
            string path = Environment.CurrentDirectory;
            string pattern = "data/UserLogo.XML";
            string[] strFileName = Directory.GetFiles(path, pattern);
            foreach (var item in strFileName)
            {
               File.Delete(item);
            }
            listView1.Items.Clear();
            load();
        }
        private void load()
        {
            user = new SqlConnect.UserList();
            int i = 0;
            foreach (DataRow dr in user.ds.Tables[0].Rows)
            {
                ListViewItem item = listView1.Items.Add(dr["UserID"].ToString());
                item.SubItems.Add((dr["Username"].ToString()).Trim());
                item.SubItems.Add((dr["UserID"].ToString()).Trim());
                if ((dr["ERP_Active"].ToString()).Trim() == "1")
                item.Checked = true;
                i++;
            }
        }
        //保存用户的Active状态
        private void saveauth()
        {
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked == true)
                {
                    user.ds.Tables[0].Rows[item.Index]["ERP_Active"] = 1;
                }
                else
                {
                    user.ds.Tables[0].Rows[item.Index]["ERP_Active"] = 0;
                }
            }
            user.Save();
        }
        #endregion


    }
}
