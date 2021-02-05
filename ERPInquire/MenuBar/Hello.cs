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
    public partial class Hello : UserControl
    {
        public string name;
        public Hello()
        {
            InitializeComponent();
        }

        private void Hello_Load(object sender, EventArgs e)
        {
            pictureBox1.BackColor = Color.Transparent;
            label1.Text =  "欢迎"+name+"使用汇源ERP数据查询系统";
        }

        private void label1_TextChanged(object sender, EventArgs e)
        {
             
        }

        private void Hello_Resize(object sender, EventArgs e)
        {
 
        }
    }
}
