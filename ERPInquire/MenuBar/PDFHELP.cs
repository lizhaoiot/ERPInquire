using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERPInquire.CustomControl
{
    public partial class PDFHELP : Form
    {
        public PDFHELP()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        }

        private void PDFHELP_Load(object sender, EventArgs e)
        {
            pdfDocument1.Load("帮助文档说明.pdf");
        }
    }
}
