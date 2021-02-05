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
    public partial class ShowBmp2 : Form
    {
        #region 初始化
        private static ShowBmp2 frm = null;
        private ShowBmp2()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }
        public static ShowBmp2 CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new ShowBmp2();
            }
            return frm;
        }
        #endregion


    }
}
