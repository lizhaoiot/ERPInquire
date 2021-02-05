using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyMIS
{
  public   class AddToolstrip
    {
      public  static void AddDTPtoToolstrip(int n, DateTimePicker dtp, ToolStrip toolStrip1)
        {
            dtp.Width =95;
            dtp.Format = DateTimePickerFormat.Custom;
            ToolStripControlHost host1 = new ToolStripControlHost(dtp);
            toolStrip1.Items.Insert(n, host1);
        }
      public static void AddLabtoToolstrip(int n, Label lable, ToolStrip toolStrip1)
        {

            lable.Width = 95;
            ToolStripControlHost host1 = new ToolStripControlHost(lable);
            toolStrip1.Items.Insert(n, host1);
        }
      public static void AddTexttoToolstrip(int n, TextBox lable, ToolStrip toolStrip1)
        {

            lable.Width = 95;
            ToolStripControlHost host1 = new ToolStripControlHost(lable);
            toolStrip1.Items.Insert(n, host1);
        }
      public static void AddspittoToolstrip(int n, ToolStrip toolStrip1)
        {
            ToolStripSeparator spit = new ToolStripSeparator();
            toolStrip1.Items.Insert(n, spit);
        }
    }
}
