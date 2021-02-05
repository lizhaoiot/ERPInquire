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
    public partial class ThinFilmTonerConversion : UserControl
    {
        #region 初始化
        public ThinFilmTonerConversion()
        {
            InitializeComponent();
        }
        #endregion

        #region 窗体设计
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

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
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

        #endregion

        #region 操作
        //计算
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            string index = comboBox3.SelectedIndex.ToString();
            switch (index)
            {
                //吨价
                case "0":
                    break;
                //令价
                case "1":
                    break;
                //张价
                case "2":
                    break;
                default:
                    break;
            }
        }
        //清除
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
        }
        //导出EXCEL
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataColumn dc = null;
            dc = dt.Columns.Add("纸张尺寸", Type.GetType("System.String"));
            dc = dt.Columns.Add("克重", Type.GetType("System.String"));
            dc = dt.Columns.Add("吨价", Type.GetType("System.String"));
            dc = dt.Columns.Add("令重", Type.GetType("System.String"));
            dc = dt.Columns.Add("令价", Type.GetType("System.String"));
            dc = dt.Columns.Add("令数", Type.GetType("System.String"));
            dc = dt.Columns.Add("张价", Type.GetType("System.String"));
            DataRow dr = dt.NewRow();
            dr["纸张尺寸"] = textBox3.Text;
         //   dr["克重"] = textBox4.Text;
            dr["吨价"] = textBox6.Text;
            dr["令重"] = textBox5.Text;
            dr["令价"] = textBox8.Text;
            dr["令数"] = textBox7.Text;
            dr["张价"] = textBox9.Text;
            dt.Rows.Add(dr);
            MyMIS.ExcleIO.OutExcel("纸价吨令转换(大度正度纸价转换、吨转令、令转吨)", dt);
        }
        #endregion
    }
}
