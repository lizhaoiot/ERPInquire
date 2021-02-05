using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Runtime.Serialization;
using System.Security;
using DocToolkit;
using Microsoft.Win32;

namespace ERPInquire.CustomControl
{
    public partial class ShowBmp1 : Form
    {
        #region 变量

        private DrawTools.DrawArea drawArea;

        #endregion

        #region 初始化
        private static ShowBmp1 frm = null;
        private ShowBmp1()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }

        public static ShowBmp1 CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new ShowBmp1();
            }
            return frm;
        }
        private void ShowBmp1_Load(object sender, EventArgs e)
        {
            //// Create draw area
            //drawArea = new DrawTools.DrawArea();
            //drawArea.Location = new Point(0, 0);
            //drawArea.Size = new Size(10, 10);
            //drawArea.Owner = this;
            //Controls.Add(drawArea);

            //// Helper objects (DocManager and others)
            //InitializeHelperObjects();

            //drawArea.Initialize(this, docManager);
            //ResizeDrawArea();

            //LoadSettingsFromRegistry();
            //drawArea.ActiveTool = DrawTools.DrawArea.DrawToolType.Rectangle;
            //drawArea.DrawFilled = false;
        }
        #endregion

    }
}
