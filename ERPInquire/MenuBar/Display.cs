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
    public partial class Display : Form
    {
        #region 定量
        bool beginMove = false;
        int currentXPosition;
        int currentYPosition;

        private static Display frm = null;
        #endregion

        #region 变量
        public DataSet ds;
        public string s1;
        #endregion

        #region 初始化
        private Display()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }

        public static Display CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new Display();
            }
            return frm;
        }
        #endregion

        #region 窗体事件
        private void Display_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            label1.Text = s1;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                this.listBox1.Items.Add(Convert.ToString(ds.Tables[0].Rows[i]["模块名字"].ToString()));
        }
        //返回
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
          //  PermissionControl pc = new PermissionControl();
          //  pc.ShowDialog();
        }
        #endregion

        #region 方法

        #endregion

        private void Display_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                beginMove = true;
                currentXPosition = MousePosition.X;//鼠标的x坐标为当前窗体左上角x坐标
                currentYPosition = MousePosition.Y;//鼠标的y坐标为当前窗体左上角y坐标
            }
        }

        private void Display_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                currentXPosition = 0; //设置初始状态
                currentYPosition = 0;
                beginMove = false;
            }
        }

        private void Display_MouseMove(object sender, MouseEventArgs e)
        {
            if (beginMove)
            {
                this.Left += MousePosition.X - currentXPosition;//根据鼠标x坐标确定窗体的左边坐标x
                this.Top += MousePosition.Y - currentYPosition;//根据鼠标的y坐标窗体的顶部，即Y坐标
                currentXPosition = MousePosition.X;
                currentYPosition = MousePosition.Y;
            }
        }
    }
}
