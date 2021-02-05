using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IO;
using SqlConnect;
using System.Data.SqlClient;
using ExcelOperation;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Com.Hui.iMRP.Utils;

namespace ERPInquire.MenuBar
{
    public partial class MaterialRequirements : Form
    {
        #region 定量
        public List<string> st1 = new List<string>();//XH
        public List<string> st4 = new List<string>();//ZBXH
        #endregion

        #region 变量
        private static string values = string.Empty;
        private static MaterialRequirements frm = null;
        #endregion

        #region 初始化
        private MaterialRequirements()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }
        public static MaterialRequirements CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new MaterialRequirements();
            }
            return frm;
        }
        private void MaterialRequirements_Load(object sender, EventArgs e)
        {
           // this.TopMost = true;
            Com.Hui.iMRP.Utils.SqlHelper.connectionStr = @"packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy";
        }

        #endregion

        #region 窗体事件
        private void MaterialRequirements_FormClosed(object sender, FormClosedEventArgs e)
        {

        }
        #endregion

        #region 公共方法
        public string ReturnValue()
        {
            return values;
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text == "") return;
            for (int i = 0; i < st1.Count; i++)
            {
                 string s1 = st1[i];
                 string s4 = st4[i];
                 string sql = "update KC_CLCK_D set CLYTRemark='" + richTextBox1.Text + "' FROM  KC_CLCK_D  where xh=" + s1 + " and zbxh="+s4+"";
                 Com.Hui.iMRP.Utils.SqlHelper.ExecCommand(sql);
                 Com.Hui.iMRP.Utils.SqlHelper.GetConnection().Close();
            }
            values = richTextBox1.Text;
            ERPInquire.MenuBar.MaterialsOfRecipients mor = ERPInquire.MenuBar.MaterialsOfRecipients.CreateInstrance();
            mor.s1();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
