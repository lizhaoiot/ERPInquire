using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IO;
using SqlConnect;
using ExcelOperation;

namespace MyMIS
{
    public partial class ReadCF2 : UserControl
    {
        DataTable dt ;
        public ReadCF2()
        {
            InitializeComponent();
            Initialization();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            toolStripTextBox1.Text = folderBrowserDialog1.SelectedPath;
            GetDieList dielist = new GetDieList();
           if( folderBrowserDialog1.SelectedPath !="")
            { 
              dielist.GetFile(folderBrowserDialog1.SelectedPath);
              foreach (CF2 s in dielist.FileList)
              {   
                  DataRow dr = dt.NewRow();
                  dr["刀线名字"] = s.Name;
                  dr["宽"] = s.Width;
                  dr["高"] = s.Height;
                  dr["模数"] = int.Parse(s.diecount.ToString());
                  dt.Rows.Add(dr); 
              }
            }
        }

        private void addDielist()
        { 
         
        
        }
        private void Initialization()
        {
            //************************************************//
            dt = new DataTable();
            dt.TableName = "Dielist";
            DataColumn dc = null;
            //dc = dt.Columns.Add("ID", Type.GetType("System.String"));
            //dc.AllowDBNull = false;
            //dc.Unique = true;
            dc = dt.Columns.Add("刀线名字", Type.GetType("System.String"));
            dc.AllowDBNull = false;
            dc.Unique = true;         
            dc = dt.Columns.Add("宽", Type.GetType("System.Double"));
            dc = dt.Columns.Add("高", Type.GetType("System.Double"));
            dc = dt.Columns.Add("模数", Type.GetType(" System.Int32"));
        
           
            //************************************************//

            //添加默认编号
            //string id = Guid.NewGuid().ToString();
            //DataRow dr = dt.NewRow();
            ////dr["ID"] = id;
            //dr["Name"] = "刀线测试";
            //dr["Width"] = 0;
            //dr["Height"] = 0;

            //dt.Rows.Add(dr);
            //************************************************//
            dataGridView1.DataSource = dt;
            dataGridView1.Refresh();
      
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
                MyMIS.ExcleIO.OutExcel("刀版打印", dt);
        }
    }
}
