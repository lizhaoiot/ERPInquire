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
    public partial class SpineCalculate : UserControl
    {
        public SpineCalculate()
        {
            InitializeComponent();
        }
        #region 输入处理

        //输入框字符处理
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
        //下拉框字符处理
        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if ((e.KeyChar == 0x2D) && (((ComboBox)sender).Text.Length == 0)) return;   //处理负数
            if (e.KeyChar > 0x20)
            {
                try
                {
                    double.Parse(((ComboBox)sender).Text + e.KeyChar.ToString());
                }
                catch
                {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }
        }
        #endregion

        #region 按钮事件
        //计算结果
        private void button1_Click(object sender, EventArgs e)
        {
            //(P数÷2)×0.001346×纸张克数 = 书脊
            if (textBox1.Text != "" && comboBox1.Text != "")
            {
                textBox2.Text =Convert.ToString((Convert.ToSingle(textBox1.Text) )/ 2 * 0.001346 * (Convert.ToSingle(comboBox1.Text)));
                textBox2.Text = textBox2.Text + "MM";
            }
        }
        //清空结果
        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            comboBox1.Text = "";
        }
        #endregion

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
    }
}
