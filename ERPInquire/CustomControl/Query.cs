using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Com.Hui.iMRP.Utils;
using System.Data.SqlClient;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HPSF;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;

namespace filesNewData
{
    public partial class Query : UserControl
    {
        #region 定量
        public string NodeName;
        #endregion

        #region 变量
        DataTable dt = new DataTable();
        #endregion

        #region 初始化
        public Query()
        {
            InitializeComponent();
            Com.Hui.iMRP.Utils.Globals.ConnectionString = @"packet size = 4096; user id = sa; pwd =; data source = 192.168.0.97; persist security info = False; initial catalog = hy";
        }
        private void Query_Load(object sender, EventArgs e)
        {
            //绑定客户名称资源
            BindCombox1();
            //绑定产品大类资源
            BindCombox2();
            //绑定产品类别资源
            BindCombox3();
            //绑定产品小类资源
            BindCombox4();

            comboBox1.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            comboBox6.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;
        }
        
        #endregion

        #region 窗体事件  
        //统计
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;//等待
                dt.Clear();
                SqlParameter[] para = {new SqlParameter("@STARTTIME", SqlDbType.DateTime),
                    new SqlParameter("@ENDTIME", SqlDbType.DateTime),
                    new SqlParameter("@KHMC", SqlDbType.VarChar),
                    new SqlParameter("@CPDL", SqlDbType.VarChar),
                    new SqlParameter("@CPLB", SqlDbType.VarChar),
                    new SqlParameter("@CPZL", SqlDbType.VarChar),
                    new SqlParameter("@CPMC", SqlDbType.VarChar),
                    new SqlParameter("@CPBM", SqlDbType.VarChar),
                    new SqlParameter("@SH", SqlDbType.Int),
                    new SqlParameter("@HX", SqlDbType.Int),
                };
                para[0].Value = dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00.000";
                para[1].Value = dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59.999";
                para[2].Value = comboBox1.Text;
                para[3].Value = comboBox3.Text;
                para[4].Value = comboBox4.Text;
                para[5].Value = comboBox5.Text;
                para[6].Value = textBox1.Text;
                para[7].Value = textBox2.Text;
                if (GetSH() == -1)
                {
                    return;
                }
                else
                {
                    para[8].Value = GetSH();
                }
                if (GetHX() == -1)
                {
                    return;
                }
                else
                {
                    para[9].Value = GetHX();
                }
                dataGridView1.DataSource = Com.Hui.iMRP.Utils.SqlHelper.ExecStoredProcedureDataTable("YUAN_CPBOM_Query", para);
                SqlHelper.GetConnection().Close();
                dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
                dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
                dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
                dataGridView1.Refresh();
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //新建
        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }
        //导出EXCEL
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }

        #endregion

        #region 公共方法
        //获得审核状态
        private int GetSH()
        {
            if (comboBox7.SelectedIndex == 0)
            {
                return 0;
            }
            else if (comboBox7.SelectedIndex == 1)
            {
                return 1;
            }
            else if (comboBox7.SelectedIndex == 2)
            {
                return 2;
            }
            else
            {
                return -1;
            }
        }
        //获得核销状态
        private int GetHX()
        {
            if (comboBox6.SelectedIndex == 0)
            {
                return 0;
            }
            else if (comboBox6.SelectedIndex == 1)
            {
                return 1;
            }
            else if (comboBox6.SelectedIndex == 2)
            {
                return 2;
            }
            else
            {
                return -1;
            }
        }
        private void Export2Excel(string nodename, DataTable dt)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = nodename + DateTime.Now.ToString("yyyyMMddhhmmss");
            dlg.Filter = "xlsx files(*.xlsx)|*.xlsx|xls files(*.xls)|*.xls|All files(*.*)|*.*";
            dlg.ShowDialog();
            if (dlg.FileName.IndexOf(":") < 0) return; //被点了"取消"
            DataTableToExcel(dt, dlg.FileName, "sheet1", true); ;
            MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public static int DataTableToExcel(DataTable dt, string fileName, string sheetName, bool isColumnWritten)
        {
            IWorkbook workbook = null;
            FileStream fileStream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0)
            {
                workbook = new XSSFWorkbook();
            }
            else if (fileName.IndexOf(".xls") > 0)
            {
                workbook = new HSSFWorkbook();
            }
            int result;
            if (workbook != null)
            {
                ISheet sheet = workbook.CreateSheet(sheetName);
                int num;
                if (isColumnWritten)
                {
                    IRow row = sheet.CreateRow(0);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        row.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                    }
                    num = 1;
                }
                else
                {
                    num = 0;
                }
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    IRow row = sheet.CreateRow(num);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if ((dt.Columns[i].ColumnName.Contains("数") || dt.Columns[i].ColumnName.Contains("价") || dt.Columns[i].ColumnName.Contains("量") || dt.Columns[i].ColumnName.Contains("面积") || dt.Columns[i].ColumnName.Contains("长") || dt.Columns[i].ColumnName.Contains("宽") || dt.Columns[i].ColumnName.Contains("宽") || dt.Columns[i].ColumnName.Contains("重") || dt.Columns[i].ColumnName.Contains("度")) && dt.Rows[j][i] != DBNull.Value)
                        {
                            try
                            {
                                row.CreateCell(i).SetCellValue(Convert.ToDouble(dt.Rows[j][i]));
                            }
                            catch
                            {
                                row.CreateCell(i).SetCellValue(dt.Rows[j][i].ToString());
                            }
                        }
                        else
                        {
                            row.CreateCell(i).SetCellValue(dt.Rows[j][i].ToString());
                        }
                    }
                    num++;
                }
                workbook.Write(fileStream);
                fileStream.Close();
                workbook.Close();
                result = num;
            }
            else
            {
                result = -1;
            }
            return result;
        }
        private void BindCombox1()
        {
            DataTable dt = new DataTable();
            string sql = "select xh,zwmc from JCZL_KHBM_M";
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "xh";
            comboBox1.DisplayMember = "zwmc";
        }
        private void BindCombox2()
        {
            DataTable dt = new DataTable();
            string sql = "select xh,MC from JCZL_CPDL";
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            comboBox3.DataSource = dt;
            comboBox3.ValueMember = "xh";
            comboBox3.DisplayMember = "MC";
        }
        private void BindCombox3()
        {
            DataTable dt = new DataTable();
            string sql = "select xh,CPLB from JCZL_CPLB";
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            comboBox4.DataSource = dt;
            comboBox4.ValueMember = "xh";
            comboBox4.DisplayMember = "CPLB";
        }
        private void BindCombox4()
        {
            DataTable dt = new DataTable();
            string sql = "select xh,XLMC from JCZL_CPLB_D";
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            comboBox5.DataSource = dt;
            comboBox5.ValueMember = "xh";
            comboBox5.DisplayMember = "XLMC";
        }

        #endregion
    }
}
