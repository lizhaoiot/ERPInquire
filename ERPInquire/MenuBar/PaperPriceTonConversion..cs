using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ERPInquire.CustomControl
{
    public partial class PaperPriceTonConversion : UserControl
    {
        public PaperPriceTonConversion()
        {
            InitializeComponent();
        }

        #region GroupBox
        //重绘GroupBox
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
        #endregion

        #region 计算方法

        #endregion

        #region 窗体事件
        private void PaperPriceTonConversion_Load(object sender, EventArgs e)
        {
            textBox10.Visible = false;
            textBox11.Visible = false;
            label13.Visible = false;
            label5.Visible = false;
            label14.Visible = false;
        }
        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
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
        //计算结果
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                string temp = comboBox1.Text;
                if (textBox1.Text !="" && comboBox1.Text !="" && comboBox2.Text!="")
               { 
                //克重
                textBox4.Text = comboBox2.Text + "克/平方米";
               
                  #region  吨价计算
            if (comboBox3.SelectedIndex == 0)
            {
                //吨价计算
                textBox6.Text = textBox1.Text + "元/吨";
                switch (temp)
                {
                    case "自定义尺寸":
                            //纸张尺寸
                            textBox3.Text = comboBox1.Text + textBox10.Text +"mm"+ "*" + textBox11.Text + "mm";
                            if (textBox10.Text !="" && textBox11.Text !="")
                          TonCalc(Convert.ToDouble(Convert.ToDouble(textBox10.Text)/1000), Convert.ToDouble( Convert.ToDouble(textBox11.Text)/1000));
                        break;
                    default:
                        //纸张尺寸
                        textBox3.Text = comboBox1.Text;
                        TonCalc(MatchStr(temp));
                        break;
                }
            }
            #endregion
               
                  #region 令价计算
            if (comboBox3.SelectedIndex ==1)
         {
                    textBox8.Text = textBox1.Text+ "元/令";
                    switch (temp)
                    {
                        case "自定义尺寸":
                            //纸张尺寸
                            textBox3.Text = comboBox1.Text + textBox10.Text + "mm" + "*" + textBox11.Text + "mm";
                            if (textBox10.Text != "" && textBox11.Text != "")
                                ZhaCalc(Convert.ToDouble(Convert.ToDouble(textBox10.Text) / 1000), Convert.ToDouble(Convert.ToDouble(textBox11.Text) / 1000));
                            break;
                        default:
                           //纸张尺寸
                           textBox3.Text = comboBox1.Text;
                            ZhaCalc(MatchStr(temp));
                           break;        
                    }
                }
            #endregion
               
                  #region 张价计算
            if (comboBox3.SelectedIndex ==2)
          {
                    textBox9.Text = textBox1.Text + "元/张";
                    switch (temp)
                    {
                        case "自定义尺寸":
                            //纸张尺寸
                            textBox3.Text = comboBox1.Text + textBox10.Text + "mm" + "*" + textBox11.Text + "mm";
                            if (textBox10.Text != "" && textBox11.Text != "")
                                TeaCalc(Convert.ToDouble(Convert.ToDouble(textBox10.Text) / 1000), Convert.ToDouble(Convert.ToDouble(textBox11.Text) / 1000));
                            break;
                        default:
                            //纸张尺寸
                            textBox3.Text = comboBox1.Text;
                            TeaCalc(MatchStr(temp));
                            break;
                        }
                }
            #endregion
                }
            }
            catch
            {
                MessageBox.Show("查询失败");
            }
        }
        //清空输入
        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "";
            comboBox2.Text = "";
            textBox1.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
        }
        //吨价计算
        private void TonCalc(string[] st)
        {
            double s1;double s2;
            s1 =Convert.ToDouble(st[0]);
            s2 = Convert.ToDouble(st[0]);
            string str1 = null;
            string str2 = null;
            string str3 = null;
            //令重
            textBox5.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2,4)) + "千克/令";
            str1 = Convert.ToString(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2);
            //令数
            textBox7.Text = Convert.ToString(Math.Round(1000 / Convert.ToSingle(str1),4)) + "令/吨";
            str2 = Convert.ToString(1000 / Convert.ToSingle(str1));
            //令价
            textBox8.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(str2),4)) + "元/令";
            str3 = Convert.ToString(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(str2));
            //单价
            textBox9.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) * Convert.ToSingle(textBox1.Text) / 1000000,4)) + "元/张";
        }
        private void TonCalc(double s1, double s2)
        {
            string str1 = null;
            string str2 = null;
            string str3 = null;
            //令重
            textBox5.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2, 4)) + "千克/令";
            str1 = Convert.ToString(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2);
            //令数
            textBox7.Text = Convert.ToString(Math.Round(1000 / Convert.ToSingle(str1), 4)) + "令/吨";
            str2 = Convert.ToString(1000 / Convert.ToSingle(str1));
            //令价
            textBox8.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(str2), 4)) + "元/令";
            str3 = Convert.ToString(Convert.ToSingle(textBox1.Text) / Convert.ToSingle(str2));
            //单价
            textBox9.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) * Convert.ToSingle(textBox1.Text) / 1000000, 4)) + "元/张";
        }
        //令价计算
        private void ZhaCalc(string[] st)
        {
            double s1; double s2;
            s1 = Convert.ToDouble(st[0]);
            s2 = Convert.ToDouble(st[0]);
            string str1 = null;
            string str2 = null;
            string str3 = null;
            //令重
            textBox5.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2, 4)) + "千克/令";
            str1 = Convert.ToString(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2);
            //令数
            textBox7.Text = Convert.ToString(Math.Round(1000 / Convert.ToSingle(str1), 4)) + "令/吨";
            str2 = Convert.ToString(1000 / Convert.ToSingle(str1));
            //吨价
            textBox6.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) * Convert.ToSingle(str2), 4)) + "元/吨";
            str3 = Convert.ToString(Convert.ToSingle(textBox1.Text)* Convert.ToSingle(str2));
            //单价
            textBox9.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) * Convert.ToSingle(str3) / 1000000, 4)) + "元/张";
        }
        private void ZhaCalc(double s1, double s2)
        {
            string str1 = null;
            string str2 = null;
            string str3 = null;
            //令重
            textBox5.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2, 4)) + "千克/令";
            str1 = Convert.ToString(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2);
            //令数
            textBox7.Text = Convert.ToString(Math.Round(1000 / Convert.ToSingle(str1), 4)) + "令/吨";
            str2 = Convert.ToString(1000 / Convert.ToSingle(str1));
            //吨价
            textBox6.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) * Convert.ToSingle(str2), 4)) + "元/吨";
            str3 = Convert.ToString(Convert.ToSingle(textBox1.Text) * Convert.ToSingle(str2));
            //单价
            textBox9.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) * Convert.ToSingle(str3) / 1000000, 4)) + "元/张";
        }
        //张价计算
        private void TeaCalc(string[] st)
        {
            double s1; double s2;
            s1 = Convert.ToDouble(st[0]);
            s2 = Convert.ToDouble(st[0]);
            string str1 = null;
            string str2 = null;
            string str3 = null;
            //令重
            textBox5.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2, 4)) + "千克/令";
            str1 = Convert.ToString(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2);
            //令数
            textBox7.Text = Convert.ToString(Math.Round(1000 / Convert.ToSingle(str1), 4)) + "令/吨";
            str2 = Convert.ToString(1000 / Convert.ToSingle(str1));
            //吨价
            textBox6.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) *1000000/Convert.ToSingle(s1 * s2) / Convert.ToSingle(comboBox2.Text), 4)) + "元/吨";
            str3 = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) * 1000000 / Convert.ToSingle(s1 * s2) / Convert.ToSingle(comboBox2.Text), 4));
            //令价
            textBox8.Text = Convert.ToString(Math.Round(Convert.ToSingle(str3) / Convert.ToSingle(str2), 4))+ "元/令";
        }
        private void TeaCalc(double s1,double s2)
        {
            string str1 = null;
            string str2 = null;
            string str3 = null;
            //令重
            textBox5.Text = Convert.ToString(Math.Round(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2, 4)) + "千克/令";
            str1 = Convert.ToString(Convert.ToSingle(s1 * s2) * Convert.ToSingle(comboBox2.Text) / 2);
            //令数
            textBox7.Text = Convert.ToString(Math.Round(1000 / Convert.ToSingle(str1), 4)) + "令/吨";
            str2 = Convert.ToString(1000 / Convert.ToSingle(str1));
            //吨价
            textBox6.Text = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) * 1000000 / Convert.ToSingle(s1 * s2) / Convert.ToSingle(comboBox2.Text), 4)) + "元/吨";
            str3 = Convert.ToString(Math.Round(Convert.ToSingle(textBox1.Text) * 1000000 / Convert.ToSingle(s1 * s2) / Convert.ToSingle(comboBox2.Text), 4));
            //令价
            textBox8.Text = Convert.ToString(Math.Round(Convert.ToSingle(str3) / Convert.ToSingle(str2), 4)) + "元/令";
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string temp = comboBox1.Text;
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
        #endregion

        #region  导出EXCEL
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
            dr["克重"] = textBox4.Text;
            dr["吨价"] = textBox6.Text;
            dr["令重"] = textBox5.Text;
            dr["令价"] = textBox8.Text;
            dr["令数"] = textBox7.Text;
            dr["张价"] = textBox9.Text;
            dt.Rows.Add(dr);
            MyMIS.ExcleIO.OutExcel("纸价吨令转换(大度正度纸价转换、吨转令、令转吨)", dt);
        }
        #endregion

        #region 取出字符
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

        #endregion

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                string temp = comboBox1.Text;
                if (textBox1.Text != "" && comboBox1.Text != "" && comboBox2.Text != "")
                {
                    //克重
                    textBox4.Text = comboBox2.Text + "克/平方米";

                    #region  吨价计算
                    if (comboBox3.SelectedIndex == 0)
                    {
                        //吨价计算
                        textBox6.Text = textBox1.Text + "元/吨";
                        switch (temp)
                        {
                            case "自定义尺寸":
                                //纸张尺寸
                                textBox3.Text = comboBox1.Text + textBox10.Text + "mm" + "*" + textBox11.Text + "mm";
                                if (textBox10.Text != "" && textBox11.Text != "")
                                    TonCalc(Convert.ToDouble(Convert.ToDouble(textBox10.Text) / 1000), Convert.ToDouble(Convert.ToDouble(textBox11.Text) / 1000));
                                break;
                            default:
                                //纸张尺寸
                                textBox3.Text = comboBox1.Text;
                                TonCalc(MatchStr(temp));
                                break;
                        }
                    }
                    #endregion

                    #region 令价计算
                    if (comboBox3.SelectedIndex == 1)
                    {
                        textBox8.Text = textBox1.Text + "元/令";
                        switch (temp)
                        {
                            case "自定义尺寸":
                                //纸张尺寸
                                textBox3.Text = comboBox1.Text + textBox10.Text + "mm" + "*" + textBox11.Text + "mm";
                                if (textBox10.Text != "" && textBox11.Text != "")
                                    ZhaCalc(Convert.ToDouble(Convert.ToDouble(textBox10.Text) / 1000), Convert.ToDouble(Convert.ToDouble(textBox11.Text) / 1000));
                                break;
                            default:
                                //纸张尺寸
                                textBox3.Text = comboBox1.Text;
                                ZhaCalc(MatchStr(temp));
                                break;
                        }
                    }
                    #endregion

                    #region 张价计算
                    if (comboBox3.SelectedIndex == 2)
                    {
                        textBox9.Text = textBox1.Text + "元/张";
                        switch (temp)
                        {
                            case "自定义尺寸":
                                //纸张尺寸
                                textBox3.Text = comboBox1.Text + textBox10.Text + "mm" + "*" + textBox11.Text + "mm";
                                if (textBox10.Text != "" && textBox11.Text != "")
                                    TeaCalc(Convert.ToDouble(Convert.ToDouble(textBox10.Text) / 1000), Convert.ToDouble(Convert.ToDouble(textBox11.Text) / 1000));
                                break;
                            default:
                                //纸张尺寸
                                textBox3.Text = comboBox1.Text;
                                TeaCalc(MatchStr(temp));
                                break;
                        }
                    }
                    #endregion
                }
            }
            catch
            {
                MessageBox.Show("查询失败");
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "";
            comboBox2.Text = "";
            textBox1.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
        }
    }
}
