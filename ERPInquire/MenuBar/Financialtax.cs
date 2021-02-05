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
using System.Data.SqlClient;
using ExcelOperation;
using System.IO;
using System.Globalization;

namespace ERPInquire.MenuBar
{
    public partial class Financialtax : Form
    {
        #region 定量

        DataTable dt = new DataTable();

        #endregion

        #region 变量

        private static Financialtax frm = null;

        #endregion

        #region 初始化
        private Financialtax()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }
        public static Financialtax CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new Financialtax();
            }
            return frm;
        }
        private void Financialtax_Load(object sender, EventArgs e)
        {
            init();
        }
        #endregion

        #region 控件事件
        //查询
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            dt.Clear();
            SqlParameter[] para ={new SqlParameter("@STARTTIME", SqlDbType.DateTime),new SqlParameter("@ENDTIME", SqlDbType.DateTime)};
            para[0].Value = dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00";
            para[1].Value = dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59";
            dt=SqlHelper.ExecStoredProcedureDataTable("YUAN_CLHS", para);
            dataGridView1.DataSource = dt;
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
            dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView1.Refresh();
            this.Cursor = Cursors.Default;//正常状态
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].ReadOnly = true;
            dataGridView1.Columns[4].ReadOnly = true;
            dataGridView1.Columns[6].ReadOnly = true;
            dataGridView1.Columns[7].ReadOnly = true;
            dataGridView1.Columns[8].ReadOnly = true;
            dataGridView1.Columns[9].ReadOnly = true;
            dataGridView1.Columns[10].ReadOnly = true;
            dataGridView1.Columns[11].ReadOnly = true;
        }
         //导出EXCEL
         private void toolStripButton2_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel("财务税率修改", SqlHelper.GetDgvToTable(dataGridView1));
        }
         private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
         {
            //if (e.RowIndex < 0) { return; }
            //DataTable dt = (dataGridView1.DataSource as DataTable);
            //if (dataGridView1.Columns[e.ColumnIndex].DataPropertyName == "发票税率")
            //{
            //    dataGridView1.Rows[e.RowIndex].Cells["库存单价"].Value = (Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["采购单价"].Value) /(1+ (Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["发票税率"].Value)/100))).ToString("f4");
            //    dataGridView1.Rows[e.RowIndex].Cells["库存金额"].Value = (Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["采购单价"].Value) / (1 + (Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["发票税率"].Value) / 100))* Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["入库数量"].Value)).ToString("f4");
            //    dataGridView1.Rows[e.RowIndex].Cells["税率金额"].Value = ((Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["采购单价"].Value) * Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["入库数量"].Value))-(Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["采购单价"].Value) / (1 + (Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["发票税率"].Value) / 100)) * Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["入库数量"].Value))).ToString("f4");
            //    dataGridView1.Rows[e.RowIndex].Cells["加税总额"].Value = (Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["采购单价"].Value) * Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["入库数量"].Value)).ToString("f4");
            //    try
            //    {
            //        string sql = "UPDATE KC_CLRK_D SET KCDJ=" + dataGridView1.Rows[e.RowIndex].Cells["库存单价"].Value.ToString() + ",KCJE=" + dataGridView1.Rows[e.RowIndex].Cells["库存金额"].Value.ToString() + ",SLJE=" + dataGridView1.Rows[e.RowIndex].Cells["税率金额"].Value.ToString() + ",JSZE=" + dataGridView1.Rows[e.RowIndex].Cells["加税总额"].Value.ToString() + ",FPSL=" + dataGridView1.Rows[e.RowIndex].Cells["发票税率"].Value.ToString() + " from KC_CLRK_M m left join KC_CLRK_D d1 on m.xh=d1.zbxh where m.BDBH= '" + dataGridView1.Rows[e.RowIndex].Cells["入库单号"].Value.ToString() + "' and d1.ZBXH_BDBH='" + dataGridView1.Rows[e.RowIndex].Cells["采购单号"].Value.ToString() + "' and d1.ZBXH_WLBM='" + dataGridView1.Rows[e.RowIndex].Cells["物料编码"].Value.ToString() + "'";
            //        SqlHelper.ExecCommand(sql);
            //        SqlHelper.GetConnection().Close();
            //    }
            //    catch
            //    {
            //        MessageBox.Show("请核对输入税率是否正确！", "警告", MessageBoxButtons.OK,MessageBoxIcon.Warning);
            //    }
            //}
        }
        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.Text == "修改")
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Selected == true)
                    {
                        dataGridView1.Rows[i].Cells["发票税率"].Value = Return1(textBox1.Text);
                        dataGridView1.Rows[i].Cells["库存单价"].Value = (Convert.ToDecimal(dataGridView1.Rows[i].Cells["采购单价"].Value) / (1 + (Convert.ToDecimal(dataGridView1.Rows[i].Cells["发票税率"].Value) / 100))).ToString("f4");
                        dataGridView1.Rows[i].Cells["库存金额"].Value = (Convert.ToDecimal(dataGridView1.Rows[i].Cells["采购单价"].Value) / (1 + (Convert.ToDecimal(dataGridView1.Rows[i].Cells["发票税率"].Value) / 100)) * Convert.ToDecimal(dataGridView1.Rows[i].Cells["入库数量"].Value)).ToString("f4");
                        dataGridView1.Rows[i].Cells["税率金额"].Value = ((Convert.ToDecimal(dataGridView1.Rows[i].Cells["采购单价"].Value) * Convert.ToDecimal(dataGridView1.Rows[i].Cells["入库数量"].Value)) - (Convert.ToDecimal(dataGridView1.Rows[i].Cells["采购单价"].Value) / (1 + (Convert.ToDecimal(dataGridView1.Rows[i].Cells["发票税率"].Value) / 100)) * Convert.ToDecimal(dataGridView1.Rows[i].Cells["入库数量"].Value))).ToString("f4");
                        dataGridView1.Rows[i].Cells["加税总额"].Value = (Convert.ToDecimal(dataGridView1.Rows[i].Cells["采购单价"].Value) * Convert.ToDecimal(dataGridView1.Rows[i].Cells["入库数量"].Value)).ToString("f4");
                        try
                        {
                            string sql = "UPDATE KC_CLRK_D SET KCDJ=" + dataGridView1.Rows[i].Cells["库存单价"].Value.ToString() + ",KCJE=" + dataGridView1.Rows[i].Cells["库存金额"].Value.ToString() + ",SLJE=" + dataGridView1.Rows[i].Cells["税率金额"].Value.ToString() + ",JSZE=" + dataGridView1.Rows[i].Cells["加税总额"].Value.ToString() + ",FPSL=" + dataGridView1.Rows[i].Cells["发票税率"].Value.ToString() + " from KC_CLRK_M m left join KC_CLRK_D d1 on m.xh=d1.zbxh where m.BDBH= '" + dataGridView1.Rows[i].Cells["入库单号"].Value.ToString() + "' and d1.ZBXH_BDBH='" + dataGridView1.Rows[i].Cells["采购单号"].Value.ToString() + "' and d1.ZBXH_WLBM='" + dataGridView1.Rows[i].Cells["物料编码"].Value.ToString() + "'";
                            SqlHelper.ExecCommand(sql);
                            SqlHelper.GetConnection().Close();
                        }
                        catch
                        {
                            MessageBox.Show("请核对输入税率是否正确！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
        }
        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Point contextMenuPoint = dataGridView1.PointToClient(Control.MousePosition);
                contextMenuStrip1.Show(dataGridView1, contextMenuPoint);
            }
        }
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("请核对输入税率是否正确！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        #endregion

        #region 公共方法
        private decimal Return1(string s)
        {
            decimal i= 0;
            if (String.IsNullOrEmpty(s) == true)
            {
                return 0;
            }
            else
            {
                return Convert.ToDecimal(s);
            }        
        }
        private void init()
        {
          //System.DateTime currentTime = DateTime.Now;
          //string year= currentTime.Year.ToString();
          //string month = currentTime.Month.ToString();
          //string dateString1 = string.Empty;
          //string dateString2 = string.Empty;
          //switch (month)
          //{
          //    case "1":
          //        dateString1 = year + "-01-" + "01";
          //        dateString2 = year + "-02-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "2":
          //        dateString1 = year + "-01-" + "26";
          //        dateString2 = year + "-02-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "3":
          //        dateString1 = year + "-02-" + "26";
          //        dateString2 = year + "-03-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "4":
          //        dateString1 = year + "-03-" + "26";
          //        dateString2 = year + "-04-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "5":
          //        dateString1 = year + "-04-" + "26";
          //        dateString2 = year + "-05-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "6":
          //        dateString1 = year + "-05-" + "26";
          //        dateString2 = year + "-06-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "7":
          //        dateString1 = year + "-06-" + "26";
          //        dateString2 = year + "-07-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "8":
          //        dateString1 = year + "-07-" + "26";
          //        dateString2 = year + "-08-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "9":
          //        dateString1 = year + "-08-" + "26";
          //        dateString2 = year + "-09-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "10":
          //        dateString1 = year + "-09-" + "26";
          //        dateString2 = year + "-10-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "11":
          //        dateString1 = year + "-10-" + "26";
          //        dateString2 = year + "-11-" + "26";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //    case "12":
          //        dateString1 = year + "-11-" + "26";
          //        dateString2 = year + "-12-" + "31";
          //        dateTimePicker1.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker1.MaxDate = Convert.ToDateTime(dateString2);
          //        dateTimePicker2.MinDate = Convert.ToDateTime(dateString1);
          //        dateTimePicker2.MaxDate = Convert.ToDateTime(dateString2);
          //        break;
          //  }
        }
        #endregion

    }
}
