using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERPInquire.CustomControl
{
    public partial class PrintingTool : UserControl
    {

        #region 初始化 
        public PrintingTool()
        {
            InitializeComponent();
        }
        private void PrintingTool_Load(object sender, EventArgs e)
        {
            this.splitContainer2.IsSplitterFixed = true;
            #region 输入初始化
            textBox10.Visible = false;
            textBox11.Visible = false;
            label13.Visible = false;
            label5.Visible = false;
            label14.Visible = false;
            textBox10.Text = "";
            textBox11.Text = "";
            textBox2.Visible = false;
            textBox1.Visible = false;
            label2.Visible = false;
            label4.Visible = false; ;
            label6.Visible = false;
            textBox2.Text = "";
            textBox1.Text = "";
            textBox3.Visible = false;
            textBox4.Visible = false;
            label8.Visible = false;
            label9.Visible = false;
            label10.Visible = false;
            textBox3.Text = "";
            textBox4.Text = "";
            radioButton1.Checked = true;
            radioButton3.Checked = true;
            radioButton10.Checked = true;
            comboBox1.Visible = false;
            #endregion 
        }
        #endregion

        #region 界面设计
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

        private void groupBox2_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(this.groupBox2.BackColor);
            e.Graphics.DrawString(this.groupBox2.Text, this.groupBox2.Font, Brushes.Black, 10, 1);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.Black, e.Graphics.MeasureString(this.groupBox2.Text, this.groupBox2.Font).Width + 8, 7, this.groupBox2.Width - 2, 7);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 1, this.groupBox2.Height - 2);
            e.Graphics.DrawLine(Pens.Black, 1, this.groupBox2.Height - 2, this.groupBox2.Width - 2, this.groupBox2.Height - 2);
            e.Graphics.DrawLine(Pens.Black, this.groupBox2.Width - 2, 7, this.groupBox2.Width - 2, this.groupBox2.Height - 2);
        }
        private void groupBox3_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(this.groupBox3.BackColor);
            e.Graphics.DrawString(this.groupBox3.Text, this.groupBox3.Font, Brushes.Black, 10, 1);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.Black, e.Graphics.MeasureString(this.groupBox3.Text, this.groupBox3.Font).Width + 8, 7, this.groupBox3.Width - 2, 7);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 1, this.groupBox3.Height - 2);
            e.Graphics.DrawLine(Pens.Black, 1, this.groupBox3.Height - 2, this.groupBox3.Width - 2, this.groupBox3.Height - 2);
            e.Graphics.DrawLine(Pens.Black, this.groupBox3.Width - 2, 7, this.groupBox3.Width - 2, this.groupBox3.Height - 2);
        }
        private void groupBox4_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(this.groupBox4.BackColor);
            e.Graphics.DrawString(this.groupBox4.Text, this.groupBox4.Font, Brushes.Black, 10, 1);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.Black, e.Graphics.MeasureString(this.groupBox4.Text, this.groupBox4.Font).Width + 8, 7, this.groupBox4.Width - 2, 7);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 1, this.groupBox4.Height - 2);
            e.Graphics.DrawLine(Pens.Black, 1, this.groupBox4.Height - 2, this.groupBox4.Width - 2, this.groupBox4.Height - 2);
            e.Graphics.DrawLine(Pens.Black, this.groupBox4.Width - 2, 7, this.groupBox4.Width - 2, this.groupBox4.Height - 2);
        }
        private void groupBox5_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(this.groupBox5.BackColor);
            e.Graphics.DrawString(this.groupBox5.Text, this.groupBox5.Font, Brushes.Black, 10, 1);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 8, 7);
            e.Graphics.DrawLine(Pens.Black, e.Graphics.MeasureString(this.groupBox5.Text, this.groupBox5.Font).Width + 8, 7, this.groupBox5.Width - 2, 7);
            e.Graphics.DrawLine(Pens.Black, 1, 7, 1, this.groupBox5.Height - 2);
            e.Graphics.DrawLine(Pens.Black, 1, this.groupBox5.Height - 2, this.groupBox5.Width - 2, this.groupBox5.Height - 2);
            e.Graphics.DrawLine(Pens.Black, this.groupBox5.Width - 2, 7, this.groupBox5.Width - 2, this.groupBox5.Height - 2);
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, panel1.ClientRectangle,
            Color.Black, 1, ButtonBorderStyle.Solid, //左边
　　　  Color.Black, 1, ButtonBorderStyle.Solid, //上边
　　　  Color.Black, 1, ButtonBorderStyle.Solid, //右边
　         Color.Black, 1, ButtonBorderStyle.Solid);//底边
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, panel1.ClientRectangle,
            Color.Black, 1, ButtonBorderStyle.Solid, //左边
            Color.Black, 1, ButtonBorderStyle.Solid, //上边
            Color.Black, 1, ButtonBorderStyle.Solid, //右边
            Color.Black, 1, ButtonBorderStyle.Solid);//底边
        }
        #endregion

        #region 数据校验
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
        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
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
        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
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
        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
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
        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
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
            }
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string temp = comboBox3.Text;
            if (temp == "自定义尺寸")
            {
                textBox2.Visible = true;
                textBox1.Visible = true;
                label2.Visible = true;
                label4.Visible = true;
                label6.Visible = true;
                textBox2.Text = "";
                textBox1.Text = "";
            }
            else
            {
                textBox2.Visible = false;
                textBox1.Visible = false;
                label2.Visible = false;
                label4.Visible = false; ;
                label6.Visible = false;
                textBox2.Text = "";
                textBox1.Text = "";
            }
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string temp = comboBox4.Text;
            if (temp == "自定义尺寸")
            {
                textBox3.Visible = true;
                textBox4.Visible = true;
                label8.Visible = true;
                label9.Visible = true;
                label10.Visible = true;
                textBox3.Text = "";
                textBox4.Text = "";
            }
            else
            {
                textBox3.Visible = false;
                textBox4.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                label10.Visible = false;
                textBox3.Text = "";
                textBox4.Text = "";
            }
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                comboBox1.Visible = true;
            }
            else
            {
                comboBox1.Visible = false;
            }
        }
        #endregion

        #region 显示图像
        //印张开版
        private void button1_Click(object sender, EventArgs e)
        {
            ShowBmp1 ShowBmp1 = ShowBmp1.CreateInstrance();
            ShowBmp1.Show();

        }
        //原纸开版
        private void button3_Click(object sender, EventArgs e)
        {
            ShowBmp2 ShowBmp2 = ShowBmp2.CreateInstrance();
            ShowBmp2.Show();
        }
        #endregion

        #region 功能栏
        //导出EXCEL
        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }
        //清除输入
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
        }
        #endregion

        #region 生成小图

        //大长方形 宽高分为  W,H
        //小长方形 宽高分为 w, h
        //横排:  a=(int)(W/w), b=(int)(H/h), a* b
        //竖排:  a=(int)(W/h), b=(int)(H/w), a* b
        //混合:  if(w>h) : 
        //m=w+h;n=w-h, a1=(int)(W/m),a2=(int)((W-a1* m)/n),长边个数 c1 = a1 + a2, 短边个数 c2=a1;
        //b1=(int)(H/m),b2=(int)((H-b1* m)/n) ，长边个数 d1 = b1 + b2, 短边个数 d2=b1;
        //c1* d2+c2* d1+c1* d1+c2* d2
        private void CreateSmallPhoto()
        {
            
        }
        #endregion

    }
}
