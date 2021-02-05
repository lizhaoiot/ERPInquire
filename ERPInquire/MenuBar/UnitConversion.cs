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
    public partial class UnitConversion : UserControl
    {
        #region 初始化
        public UnitConversion()
        {
            InitializeComponent();
        }
        private void Unit_conversion_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;
        }
        #endregion

        //清除输入
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox4.Text = "";
            textBox6.Text = "";
        }
        #region 数据校验
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
        #endregion

        #region 方法
        //计算
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            //面积单位
            if (textBox2.Text != "")
            {
                string index1 = comboBox2.SelectedIndex.ToString();
                string index2 = comboBox1.SelectedIndex.ToString();
                switch (index1)
                {
                    //平方米
                    case "0":
                        switch (index2)
                        {
                            //平方米
                            case "0":
                                textBox1.Text = textBox2.Text;
                                break;
                            //平方分米
                            case "1":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text)*100);
                                break;
                            //平方厘米
                            case "2":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) * 10000);
                                break;
                            //平方毫米
                            case "3":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) * 1000000);
                                break;
                            default:
                                break;
                        }
                        break;
                    //平方分米
                    case "1":
                        switch (index2)
                        {
                            //平方米
                            case "0":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text)/ 100);
                                break;
                            //平方分米
                            case "1":
                                textBox1.Text = textBox2.Text;
                                break;
                            //平方厘米
                            case "2":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) * 100);
                                break;
                            //平方毫米
                            case "3":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) * 10000);
                                break;
                            default:
                                break;
                        }
                        break;
                     //平方厘米
                    case "2":
                        switch (index2)
                        {
                            //平方米
                            case "0":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) / 10000);
                                break;
                            //平方分米
                            case "1":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) / 100);
                                break;
                            //平方厘米
                            case "2":
                                textBox1.Text = textBox2.Text;
                                break;
                            //平方毫米
                            case "3":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) * 100);
                                break;
                            default:
                                break;
                        }
                        break;
                     //平方毫米
                    case "3":
                        switch (index2)
                        {
                            //平方米
                            case "0":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) / 1000000);
                                break;
                            //平方分米
                            case "1":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) / 10000);
                                break;
                            //平方厘米
                            case "2":
                                textBox1.Text = Convert.ToString(Convert.ToSingle(textBox2.Text) / 100);
                                break;
                            //平方毫米
                            case "3":
                                textBox1.Text = textBox2.Text;
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        break;
                }
            }
            //长度单位
            if (textBox4.Text != "")
            {
                string index1 = comboBox3.SelectedIndex.ToString();
                string index2 = comboBox4.SelectedIndex.ToString();
                switch (index1)
                {
                    //米
                    case "0":
                        switch (index2)
                        {
                            //米
                            case "0":
                                textBox3.Text = textBox4.Text;
                                break;
                            //分米
                            case "1":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) *10);
                                break;
                            //厘米
                            case "2":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) * 100);
                                break;
                            //毫米
                            case "3":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) * 1000);
                                break;
                            default:
                                break;
                        }
                        break;
                    //分米
                    case "1":
                        switch (index2)
                        {
                            //米
                            case "0":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) / 10);
                                break;
                            //分米
                            case "1":
                                textBox3.Text = textBox4.Text;
                                break;
                            //厘米
                            case "2":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) *10);
                                break;
                            //毫米
                            case "3":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) * 100);
                                break;
                            default:
                                break;
                        }
                        break;
                    //厘米
                    case "2":
                        switch (index2)
                        {
                            //米
                            case "0":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) / 100);
                                break;
                            //分米
                            case "1":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text)/10);
                                break;
                            //厘米
                            case "2":
                                textBox3.Text = textBox4.Text;
                                break;
                            //毫米
                            case "3":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text)* 10);
                                break;
                            default:
                                break;
                        }
                        break;
                    //毫米
                    case "3":
                        switch (index2)
                        {
                            //米
                            case "0":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) / 1000);
                                break;
                            //分米
                            case "1":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) / 100);
                                break;
                            //厘米
                            case "2":
                                textBox3.Text = Convert.ToString(Convert.ToSingle(textBox4.Text) / 10);
                                break;
                            //毫米
                            case "3":
                                textBox3.Text = textBox4.Text;
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        break;
                }
            }
            //重量单位
            if (textBox6.Text != "")
            {
                string index1 = comboBox5.SelectedIndex.ToString();
                string index2 = comboBox6.SelectedIndex.ToString();
                switch (index1)
                {
                    //吨
                    case "0":
                        switch (index2)
                        {
                            //吨
                            case "0":
                                textBox5.Text = textBox6.Text;
                                break;
                            //公斤
                            case "1":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) *1000);
                                break;
                            //千克
                            case "2":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) * 1000);
                                break;
                            //克
                            case "3":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) * 1000000);
                                break;
                            //毫克
                            case "4":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) * 1000000000);
                                break;
                            default:
                                break;
                        }
                        break;
                    //公斤
                    case "1":
                        switch (index2)
                        {
                            //吨
                            case "0":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) /1000);
                                break;
                            //公斤
                            case "1":
                                textBox5.Text = textBox6.Text;
                                break;
                            //千克
                            case "2":
                                textBox5.Text = textBox6.Text;
                                break;
                            //克
                            case "3":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) *1000);
                                break;
                            //毫克
                            case "4":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) * 1000000);
                                break;
                            default:
                                break;
                        }
                        break;
                    //千克
                    case "2":
                        switch (index2)
                        {
                            //吨
                            case "0":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) / 1000);
                                break;
                            //公斤
                            case "1":
                                textBox5.Text = textBox6.Text;
                                break;
                            //千克
                            case "2":
                                textBox5.Text = textBox6.Text;
                                break;
                            //克
                            case "3":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) * 1000);
                                break;
                            //毫克
                            case "4":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) * 1000000);
                                break;
                            default:
                                break;
                        }
                        break;
                    //克
                    case "3":
                        switch (index2)
                        {
                            //吨
                            case "0":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) / 1000000);
                                break;
                            //公斤
                            case "1":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) / 1000);
                                break;
                            //千克
                            case "2":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) /1000);
                                break;
                            //克
                            case "3":
                                textBox5.Text = textBox6.Text;
                                break;
                            //毫克
                            case "4":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) *1000);
                                break;
                            default:
                                break;
                        }
                        break;
                    //毫克
                    case "4":
                        switch (index2)
                        {
                            //吨
                            case "0":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) / 1000000000);
                                break;
                            //公斤
                            case "1":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) / 1000000);
                                break;
                            //千克
                            case "2":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) / 1000000);
                                break;
                            //克
                            case "3":
                                textBox5.Text = Convert.ToString(Convert.ToSingle(textBox6.Text) /1000);
                                break;
                            //毫克
                            case "4":
                                textBox5.Text = textBox6.Text;
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        break;
                }
            }
        }
        #endregion

    }
}
