using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using ExcelOperation;

namespace ERPInquire.CustomControl
{
    public partial class ConversionProductOriginalPaper : UserControl
    {
        #region 变量

        DataTable dt = new DataTable();

        #endregion

        #region 初始化
        public ConversionProductOriginalPaper()
        {
            InitializeComponent();
        }
        private void ConversionProductOriginalPaper_Load(object sender, EventArgs e)
        {
            tableLayoutPanel1.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(tableLayoutPanel1, true, null);
            init();
            textBox10.Visible = false;
            textBox11.Visible = false;
            label13.Visible = false;
            label5.Visible = false;
            label14.Visible = false;
            textBox10.Text = "";
            textBox11.Text = "";
            DataColumn dc = null;
            dc = dt.Columns.Add("成品数量", Type.GetType("System.String"));
            dc = dt.Columns.Add("印张数量", Type.GetType("System.String"));
            dc = dt.Columns.Add("原纸数量", Type.GetType("System.String"));
            dc = dt.Columns.Add("原纸面积", Type.GetType("System.String"));
            dc = dt.Columns.Add("原纸重量", Type.GetType("System.String"));
            dc = dt.Columns.Add("原纸规格", Type.GetType("System.String"));
            dc = dt.Columns.Add("克重", Type.GetType("System.String"));
            dc = dt.Columns.Add("得率", Type.GetType("System.String"));
            dc = dt.Columns.Add("开数", Type.GetType("System.String"));
            textBox1.Focus();
            
        }
        private void init()
        {
            this.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(this, true, null);
            Label lb1 = new Label();
            Label lb2 = new Label();
            Label lb3 = new Label();
            Label lb4 = new Label();
            Label lb5 = new Label();
            Label lb6 = new Label();
            Label lb7 = new Label();
            Label lb8 = new Label();
            Label lb9 = new Label();
            Label lb10 = new Label();
            Label lb11 = new Label();
            lb1.Text = "名称";
            lb1.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb1, 0, 0);
            lb2.Text = "数值";
            lb2.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb2, 0, 1);
            lb3.Text = "成品数量";
            lb3.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb3, 1, 0);
            lb4.Text = "印张数量";
            lb4.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb4, 2, 0);
            lb5.Text = "原纸数量";
            lb5.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb5, 3, 0);
            lb6.Text = "原纸面积";
            lb6.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb6, 4, 0);
            lb7.Text = "原纸重量";
            lb7.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb7, 5, 0);
            lb8.Text = "原纸规格";
            lb8.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb8, 6, 0);
            lb9.Text = "克重";
            lb9.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb9, 7, 0);
            lb10.Text = "得率";
            lb10.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb10, 8, 0);
            lb11.Text = "开数";
            lb11.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb11, 9, 0);
        }
        #endregion

        #region 数据校验
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if ((e.KeyChar == 0x2D) && (((TextBox)sender).Text.Length == 0)) return;   //处理负数
            if (e.KeyChar > 0x20)
            {
                try
                {
                    double.Parse(((TextBox)sender).Text + e.KeyChar.ToString());
                }
                catch
                {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }
        }
        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if ((e.KeyChar == 0x2D) && (((TextBox)sender).Text.Length == 0)) return;   //处理负数
            if (e.KeyChar > 0x20)
            {
                try
                {
                    double.Parse(((TextBox)sender).Text + e.KeyChar.ToString());
                }
                catch
                {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }
        }
        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if ((e.KeyChar == 0x2D) && (((TextBox)sender).Text.Length == 0)) return;   //处理负数
            if (e.KeyChar > 0x20)
            {
                try
                {
                    double.Parse(((TextBox)sender).Text + e.KeyChar.ToString());
                }
                catch
                {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if ((e.KeyChar == 0x2D) && (((TextBox)sender).Text.Length == 0)) return;   //处理负数
            if (e.KeyChar > 0x20)
            {
                try
                {
                    double.Parse(((TextBox)sender).Text + e.KeyChar.ToString());
                }
                catch
                {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }
        }
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if ((e.KeyChar == 0x2D) && (((TextBox)sender).Text.Length == 0)) return;   //处理负数
            if (e.KeyChar > 0x20)
            {
                try
                {
                    double.Parse(((TextBox)sender).Text + e.KeyChar.ToString());
                }
                catch
                {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }
        }
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if ((e.KeyChar == 0x2D) && (((TextBox)sender).Text.Length == 0)) return;   //处理负数
            if (e.KeyChar > 0x20)
            {
                try
                {
                    double.Parse(((TextBox)sender).Text + e.KeyChar.ToString());
                }
                catch
                {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }
        }
        #endregion

        #region 窗体设计

        private void groupBox1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(this.groupBox1.BackColor);
            e.Graphics.DrawString(this.groupBox1.Text, this.groupBox1.Font, Brushes.Black, 10, 1);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.Black, e.Graphics.MeasureString(this.groupBox1.Text, this.groupBox1.Font).Width + 8, 7, this.groupBox1.Width - 2, 7);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 1, this.groupBox1.Height - 2);
            e.Graphics.DrawLine(Pens.Black, 1, this.groupBox1.Height - 2, this.groupBox1.Width - 2, this.groupBox1.Height - 2);
            e.Graphics.DrawLine(Pens.Black, this.groupBox1.Width - 2, 7, this.groupBox1.Width - 2, this.groupBox1.Height - 2);
        }


        private void tableLayoutPanel1_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
        {
            // 防止闪屏  
            this.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(this, true, null);
            Pen pp = new Pen(BorderColor);
            e.Graphics.DrawRectangle(pp, e.CellBounds.X, e.CellBounds.Y, e.CellBounds.X + this.panel1.Width - 1, e.CellBounds.Y + this.panel1.Height-1);
        }
        private Color borderColor = Color.Black;
        public Color BorderColor
        {
            get { return borderColor; }
            set { borderColor = value; }
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string temp = comboBox2.Text;
            if (temp == "自定义尺寸")
            {
                textBox10.Visible = true;
                textBox11.Visible = true;
                label13.Visible = true;
                label5.Visible = true;
                label14.Visible = true;
                textBox10.Text = "";
                textBox11.Text = "";
                textBox1.Focus();
            }
            else
            {
                textBox10.Visible = false;
                textBox11.Visible = false;
                label13.Visible = false;
                label5.Visible = false;
                label14.Visible = false;
                textBox10.Text = "";
                textBox11.Text = "";
                textBox1.Focus();
            }
        }
        #endregion

        #region 计算
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            dt.Clear();
            tableLayoutPanel1.Controls.Clear();
            init();
            if (textBox1.Text == "") return;
            if (textBox2.Text == "") return;
            if (textBox3.Text == "") return;
            if (textBox4.Text == "") return;
            #region 初始化
            string str1 = null;
            string str2 = null;
            string str3 = null;
            string[] str = null;
            Label lb1 = new Label();
            Label lb2 = new Label();
            Label lb3 = new Label();
            Label lb4 = new Label();
            Label lb5 = new Label();
            Label lb6 = new Label();
            Label lb7 = new Label();
            Label lb8 = new Label();
            Label lb9 = new Label();
            int index = comboBox1.SelectedIndex;
            string temp = comboBox2.Text;
            //克重
            lb7.Text = Convert.ToString(textBox2.Text)+ "克/平方米";
            lb7.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb7, 7, 1);
            //得率
            lb8.Text = Convert.ToString(textBox3.Text);
            lb8.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb8, 8, 1);
            //开数
            lb9.Text = Convert.ToString(textBox4.Text);
            lb9.Anchor = AnchorStyles.None;
            this.tableLayoutPanel1.Controls.Add(lb9, 9, 1);
            #endregion
            switch (index)
            {
                #region  成品数量
                case 0:
                    if (temp == "自定义尺寸")
                    {
                        if (textBox10.Text == "") return;
                        if (textBox11.Text == "") return;
                        //成品数量
                        lb1.Text = textBox1.Text+"张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1, 1, 1);
                        //印张数量
                        lb2.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(textBox3.Text))) + "张";
                        str1 = Convert.ToString(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(textBox3.Text));
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2, 2, 1);
                        //原纸数量
                        lb3.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str1) / Convert.ToSingle(textBox4.Text))) + "张";
                        str2 = Convert.ToString(Convert.ToSingle(str1) / Convert.ToSingle(textBox4.Text));
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox10.Text) / 1000 * Convert.ToSingle(textBox11.Text) / 1000 * Convert.ToSingle(str2),4))+ "平方米";
                        str3 = Convert.ToString(Convert.ToSingle(textBox10.Text) / 1000 * Convert.ToSingle(textBox11.Text) / 1000 * Convert.ToSingle(str2));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4, 4, 1);
                        //原纸重量
                        lb5.Text = Convert.ToString(Math.Round(Convert.ToSingle(str3) * Convert.ToSingle(textBox2.Text) / 1000000,4))+"吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(textBox10.Text+"*"+textBox11.Text);
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6, 6, 1);
                    }
                    else
                    {
                        str=MatchStr(comboBox2.Text);
                        //成品数量
                        lb1.Text = textBox1.Text + "张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1,1,1);
                        //印张数量
                        lb2.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text)/Convert.ToSingle(textBox3.Text))) + "张";
                        str1= Convert.ToString(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(textBox3.Text));
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2,2,1);
                        //原纸数量
                        lb3.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str1)/Convert.ToSingle(textBox4.Text))) + "张";
                        str2= Convert.ToString(Convert.ToSingle(str1) / Convert.ToSingle(textBox4.Text));
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(str[0])/1000*Convert.ToSingle(str[1])/1000*Convert.ToSingle(str2),4)) + "平方米";
                        str3= Convert.ToString(Convert.ToSingle(str[0]) / 1000 * Convert.ToSingle(str[1]) / 1000 * Convert.ToSingle(str2));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4,4,1);
                        //原纸重量
                        lb5.Text = Convert.ToString(Math.Round(Convert.ToSingle(str3)*Convert.ToSingle(textBox2.Text)/1000000,4)) + "吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(Guige(comboBox2.Text));
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6,6, 1);
                    }
                    break;
                #endregion

                #region  印张数量
                case 1:
                    if (temp == "自定义尺寸")
                    {
                        if (textBox10.Text == "") return;
                        if (textBox11.Text == "") return;
                        //成品数量
                        lb1.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text) * Convert.ToSingle(textBox3.Text))) + "张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1, 1, 1);
                        //印张数量
                        lb2.Text = textBox1.Text + "张";
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2, 2, 1);
                        //原纸数量
                        lb3.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(textBox4.Text))) + "张";
                        str1 = Convert.ToString(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(textBox4.Text));
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox10.Text) / 1000 * Convert.ToSingle(textBox11.Text) / 1000 * Convert.ToSingle(str1),4)) + "平方米";
                        str2 = Convert.ToString(Convert.ToSingle(textBox10.Text) / 1000 * Convert.ToSingle(textBox11.Text) / 1000 * Convert.ToSingle(str1));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4, 4, 1);
                        //原纸重量
                        lb5.Text = Convert.ToString(Math.Round(Convert.ToSingle(str2) * Convert.ToSingle(textBox2.Text) / 1000000,4)) + "吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(textBox10.Text + "*" + textBox11.Text);
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6, 6, 1);
                    }
                    else
                    {
                        str = MatchStr(comboBox2.Text);
                        //成品数量
                        lb1.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text)*Convert.ToSingle(textBox3.Text))) + "张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1, 1, 1);
                        //印张数量
                        lb2.Text = textBox1.Text + "张";
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2, 2, 1);
                        //原纸数量
                        lb3.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text)/Convert.ToSingle(textBox4.Text))) + "张";
                        str1= Convert.ToString(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(textBox4.Text));
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(str[0]) / 1000 * Convert.ToSingle(str[1]) / 1000 * Convert.ToSingle(str1),4)) + "平方米";
                        str2 = Convert.ToString(Convert.ToSingle(str[0]) / 1000 * Convert.ToSingle(str[1]) / 1000 * Convert.ToSingle(str1));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4, 4, 1);
                        //原纸重量
                        lb5.Text = Convert.ToString(Math.Round(Convert.ToSingle(str2) * Convert.ToSingle(textBox2.Text) / 1000000,4)) + "吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(Guige(comboBox2.Text));
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6, 6, 1);
                    }
                    break;
                #endregion

                #region  原纸数量
                case 2:
                    if (temp == "自定义尺寸")
                    {
                        if (textBox10.Text == "") return;
                        if (textBox11.Text == "") return;
                        //印张数量
                        lb2.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text) * Convert.ToSingle(textBox4.Text))) + "张";
                        str1 = Convert.ToString(Convert.ToSingle(textBox1.Text) * Convert.ToSingle(textBox4.Text));
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2, 2, 1);
                        //原纸数量
                        lb3.Text = textBox1.Text + "张";
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //成品数量
                        lb1.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str1) * Convert.ToSingle(textBox3.Text))) + "张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1, 1, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox10.Text) / 1000 * Convert.ToSingle(textBox11.Text) / 1000 * Convert.ToSingle(str1),4)) + "平方米";
                        str2 = Convert.ToString(Convert.ToSingle(textBox10.Text) / 1000 * Convert.ToSingle(textBox11.Text) / 1000 * Convert.ToSingle(str1));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4, 4, 1);
                        //原纸重量
                        lb5.Text = Convert.ToString(Math.Round(Convert.ToSingle(str2) * Convert.ToSingle(textBox2.Text) / 1000000,4)) + "吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(textBox10.Text + "*" + textBox11.Text);
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6, 6, 1);
                    }
                    else
                    {
                        str = MatchStr(comboBox2.Text);
                        //印张数量
                        lb2.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(textBox1.Text)*Convert.ToSingle(textBox4.Text))) + "张";
                        str1 = Convert.ToString(Convert.ToSingle(textBox1.Text) * Convert.ToSingle(textBox4.Text));
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2, 2, 1);
                        //原纸数量
                        lb3.Text = textBox1.Text + "张";
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //成品数量
                        lb1.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str1) * Convert.ToSingle(textBox3.Text))) + "张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1, 1, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(str[0]) / 1000 * Convert.ToSingle(str[1]) / 1000 * Convert.ToSingle(str1),4)) + "平方米";
                        str2 = Convert.ToString(Convert.ToSingle(str[0]) / 1000 * Convert.ToSingle(str[1]) / 1000 * Convert.ToSingle(str1));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4, 4, 1);
                        //原纸重量
                        lb5.Text = Convert.ToString(Math.Round(Convert.ToSingle(str2) * Convert.ToSingle(textBox2.Text) / 1000000,4)) + "吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(Guige(comboBox2.Text));
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6, 6, 1);
                    }
                    break;
                #endregion

                #region 原纸吨数
                case 3:
                    if (temp == "自定义尺寸")
                    {
                        if (textBox10.Text == "") return;
                        if (textBox11.Text == "") return;
                        //原纸重量
                        lb5.Text = textBox1.Text + "吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) * 1000000 / Convert.ToSingle(textBox2.Text),4)) + "平方米";
                        str1 = Convert.ToString(Convert.ToSingle(textBox1.Text) * 1000000 / Convert.ToSingle(textBox2.Text));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4, 4, 1);
                        //原纸数量
                        lb3.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str1) / Convert.ToSingle(textBox10.Text) * 1000 / Convert.ToSingle(textBox11.Text) * 1000)) + "张";
                        str2 = Convert.ToString(Convert.ToSingle(str1) / Convert.ToSingle(textBox10.Text) * 1000 / Convert.ToSingle(textBox11.Text) * 1000);
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //印张数量
                        lb2.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str2) * Convert.ToSingle(textBox4.Text))) + "张";
                        str3 = Convert.ToString(Convert.ToSingle(str2) * Convert.ToSingle(textBox4.Text));
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2, 2, 1);
                        //成品数量
                        lb1.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str3) * Convert.ToSingle(textBox3.Text))) + "张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1, 1, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(textBox10.Text + "*" + textBox11.Text);
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6, 6, 1);
                    }
                    else
                    {
                        str = MatchStr(comboBox2.Text);
                        //原纸重量
                        lb5.Text = textBox1.Text + "吨";
                        lb5.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb5, 5, 1);
                        //原纸面积
                        lb4.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) *1000000/Convert.ToSingle(textBox2.Text),4)) + "平方米";
                        str1= Convert.ToString(Convert.ToSingle(textBox1.Text) * 1000000 / Convert.ToSingle(textBox2.Text));
                        lb4.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb4, 4, 1);
                        //原纸数量
                        lb3.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str1)/Convert.ToSingle(str[0])*1000/Convert.ToSingle(str[1])*1000)) + "张";
                        str2 = Convert.ToString(Convert.ToSingle(str1) / Convert.ToSingle(str[0]) * 1000 / Convert.ToSingle(str[1]) * 1000);
                        lb3.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb3, 3, 1);
                        //印张数量
                        lb2.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str2) * Convert.ToSingle(textBox4.Text))) + "张";
                        str3 = Convert.ToString(Convert.ToSingle(str2) * Convert.ToSingle(textBox4.Text));
                        lb2.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb2, 2, 1);
                        //成品数量
                        lb1.Text = Convert.ToString(Math.Ceiling(Convert.ToSingle(str3) * Convert.ToSingle(textBox3.Text))) + "张";
                        lb1.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb1, 1, 1);
                        //原纸规格
                        lb6.Text = Convert.ToString(Guige(comboBox2.Text));
                        lb6.Anchor = AnchorStyles.None;
                        this.tableLayoutPanel1.Controls.Add(lb6, 6, 1);
                    }
                    break;
                    #endregion
            }
            string[] sdt = { lb1.Text,lb2.Text,lb3.Text,lb4.Text,lb5.Text,lb6.Text,lb7.Text,lb8.Text,lb9.Text};
            CreateTable(sdt);
        }

        #endregion

        #region 清除输入
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox1.Focus();
        }

        #endregion

        #region 导出EXCEL

        //根据查询结果生成Datatable
        private void CreateTable(string[] s)
        {
            DataRow dr = dt.NewRow();
            dr["成品数量"] = s[0];
            dr["印张数量"] = s[1];
            dr["原纸数量"] = s[2];
            dr["原纸面积"] = s[3];
            dr["原纸重量"] = s[4];
            dr["原纸规格"] = s[5];
            dr["克重"] = s[6];
            dr["得率"] = s[7];
            dr["开数"] = s[8];
            dt.Rows.Add(dr);
            dataGridView1.DataSource = dt;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
                ExcelOperByInterop a = new ExcelOperByInterop();
                a.PutOutExcelByDataGridView("产品与原纸换算", dataGridView1, false);
        }
        #endregion

        #region 方法
        private string[] MatchStr(string str)
        {
            string[] st = null;
            string pattern = @"\(.*?\)";//匹配模式
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            MatchCollection matches = regex.Matches(str);
            StringBuilder sb = new StringBuilder();
            foreach (Match match in matches)
            {
                string value = match.Value.Trim('(', ')');
                sb.AppendLine(value);
                string temp = sb.ToString();
                st = temp.Split('*');
            }
            return st;
        }
        private string Guige(string str)
        {
            string temp = null;
            string pattern = @"\(.*?\)";//匹配模式
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            MatchCollection matches = regex.Matches(str);
            StringBuilder sb = new StringBuilder();
            foreach (Match match in matches)
            {
                string value = match.Value.Trim('(', ')');
                sb.AppendLine(value);
                temp = sb.ToString();
            }
            return temp;
        }
        #endregion
    }
}
