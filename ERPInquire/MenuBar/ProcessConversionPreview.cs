using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ERPInquire.MenuBar
{
    public partial class ProcessConversionPreview : Form
    {
        #region 定量
        public static DataTable dt1 = null;
        public static DataTable dt2 = null;
        public static DataTable dt3 = null;
        public static DataTable dt4 = null;
        #endregion

        #region 变量

        #endregion

        #region 初始化
        public ProcessConversionPreview()
        {
            InitializeComponent();
        }
        private void ProcessConversionPreview_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            init1();
            init2();
        }
        private void init1()
        {
            dataGridView4.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView4.Refresh();
            dataGridView4.DataSource = dt1;
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                if (dataGridView4.Rows[i].Cells["工序编码"].Value.ToString() == "" && dataGridView4.Rows[i].Cells["工序类别"].Value.ToString() == "" && dataGridView4.Rows[i].Cells["工序名称"].Value.ToString() == "")
                {
                    this.dataGridView4.Rows[i].Selected = false;
                    this.dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView4.Rows[i].Cells["产品名称"].Value = dataGridView4.Rows[i].Cells["产品名称"].Value + "该产品档案处于为审核状态";
                }
            }
        }
        private void init2()
        {
            dataGridView2.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView2.Refresh();
            dataGridView2.DataSource = dt3;
            for (int ji = 0; ji < dataGridView2.Rows.Count; ji++)
            {
                if (dataGridView2.Rows[ji].Cells["物料大类"].Value.ToString() == "" && dataGridView2.Rows[ji].Cells["物料子类"].Value.ToString() == "" && dataGridView2.Rows[ji].Cells["物料编码"].Value.ToString() == "" && dataGridView2.Rows[ji].Cells["物料名称"].Value.ToString() == "")
                {
                    this.dataGridView2.Rows[ji].Selected = false;
                    this.dataGridView2.Rows[ji].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView2.Rows[ji].Cells["产品名称"].Value = dataGridView2.Rows[ji].Cells["产品名称"].Value + "该产品档案处于为审核状态";
                }
            }
        }
        private void init3()
        {
            dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt2;
        }
        private void init4()
        {
            dataGridView3.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView3.Refresh();
            dataGridView3.DataSource = dt4;
        }
        #endregion

        #region 窗体事件

        #endregion

        #region 方法

        #endregion

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = dt3;
            for (int ji = 0; ji < dataGridView2.Rows.Count; ji++)
            {
                if (dataGridView2.Rows[ji].Cells["物料大类"].Value.ToString() == "" && dataGridView2.Rows[ji].Cells["物料子类"].Value.ToString() == "" && dataGridView2.Rows[ji].Cells["物料编码"].Value.ToString() == "" && dataGridView2.Rows[ji].Cells["物料名称"].Value.ToString() == "")
                {
                    this.dataGridView2.Rows[ji].Selected = false;
                    this.dataGridView2.Rows[ji].DefaultCellStyle.BackColor = Color.Red;
                    Regex reg = new Regex("该产品档案处于为审核状态");
                    Match m = reg.Match(dataGridView2.Rows[ji].Cells["产品名称"].Value.ToString());
                    if (!m.Success)
                        dataGridView2.Rows[ji].Cells["产品名称"].Value = dataGridView2.Rows[ji].Cells["产品名称"].Value + "该产品档案处于为审核状态";
                }
            }
            dataGridView4.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView4.Refresh();
            dataGridView4.DataSource = dt1;
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                if (dataGridView4.Rows[i].Cells["工序编码"].Value.ToString() == "" && dataGridView4.Rows[i].Cells["工序类别"].Value.ToString() == "" && dataGridView4.Rows[i].Cells["工序名称"].Value.ToString() == "")
                {
                    this.dataGridView4.Rows[i].Selected = false;
                    this.dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    Regex reg = new Regex("该产品档案处于为审核状态");
                    Match m = reg.Match(dataGridView4.Rows[i].Cells["产品名称"].Value.ToString());
                    if(!m.Success)
                      dataGridView4.Rows[i].Cells["产品名称"].Value = dataGridView4.Rows[i].Cells["产品名称"].Value + "该产品档案处于为审核状态";
                }
            }
            dataGridView3.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView3.Refresh();
            dataGridView3.DataSource = dt4;
            dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt2;
        }
        #region 未审核信息
          private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            List<string> lst = new List<string>();
            for (int i = 0; i < dataGridView4.RowCount; i++)
            {
                if (dataGridView4.Rows[i].DefaultCellStyle.BackColor == Color.Red)
                {
                    lst.Add(dataGridView4.Rows[i].Cells["产品编码"].Value.ToString());
                }
            }
            MyMIS.ExcleIO.Export2Excel1("未审核的产品编码", ToDataTable(lst));
        }
          //List转换成DataTable
          private DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("未审核的产品编码", Type.GetType("System.String"));
            dt.Columns.Add(dc1);
            //以上代码完成了DataTable的构架，但是里面是没有任何数据的
            for (int i = 0; i < items.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr["未审核的产品编码"] = items[i];
                dt.Rows.Add(dr);
            }
            return dt;
        }
        #endregion
    }
}
