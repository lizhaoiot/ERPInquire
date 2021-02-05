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
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ERPInquire.CustomControl
{
    public partial class CommonControl2 : UserControl
    {
        #region 定量

        public  string ProName;
        public  string TheKeyword;
        public string NodeName;
        public string TimeKey;
        #endregion

        #region 变量

        DateTimePicker sdate;
        DateTimePicker edate;
        public string sqlstr;
        public ConnectStr connect;
        DataTable dt = new DataTable();
        #endregion

        #region 初始化
        public CommonControl2()
        {
            InitializeComponent();
            Initialize();
            this.dataGridView1.RowsAdded += new DataGridViewRowsAddedEventHandler(dataGridView1_RowsAdded);
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行被选中
        }

        private void Initialize()
        {
            DateTime dt = DateTime.Today;
            this.sdate = new DateTimePicker();
            this.edate = new DateTimePicker();
            this.sdate.Value = dt.AddDays(-2d);
            this.edate.Value = dt;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(4, sdate, toolStrip1);
            MyMIS.AddToolstrip.AddDTPtoToolstrip(6, edate, toolStrip1);
        }

        #endregion

        #region 窗体事件
        private void CommonControl2_Load(object sender, EventArgs e)
        {
            toolStripLabel2.Text = TheKeyword;
            toolStripLabel4.Text = TimeKey;
        }
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < e.RowCount; i++)
                this.dataGridView1.Rows[e.RowIndex + i].HeaderCell.Value = (e.RowIndex + i + 1).ToString();
        }
        //查询
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            //实现首次查询
            if (toolStripTextBox2.Text.Equals(""))
            {
                dt.Clear();
                SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
                string s = connect.GetConnectStr("ERP");
                string connectionString = s;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = ProName;
                        SqlParameter[] para ={
                                     new SqlParameter("@DAH",SqlDbType.VarChar),
                                     new SqlParameter("@STARTTIME",SqlDbType.DateTime),
                                     new SqlParameter("@ENDTIME",SqlDbType.DateTime)
              };
                        para[0].Value = toolStripTextBox1.Text;
                        para[1].Value = sdate.Value.ToString("yyyy-MM-dd") + " 00:00:00";
                        para[2].Value = edate.Value.ToString("yyyy-MM-dd") + " 23:59:59";
                        try
                        {
                            cmd.CommandTimeout = 60 * 60 * 10000;
                            cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(dt);
                            conn.Close();
                            dataGridView1.DataSource = dt;
                            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
                            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
                            toolStripLabel3.Text = "总行数:" + (dt.Rows.Count).ToString();
                            dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
                            dataGridView1.Refresh();
                            //获得datatable的列名
                            List<string> ls =SqlHelper.GetColumnsByDataTable(dt);
                            toolStripComboBox1.ComboBox.DataSource = ls;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
            //实现再查询
            else
            {
                try
                {
                //获得现在所选列的列索引
                int i = toolStripComboBox1.ComboBox.SelectedIndex;
                string dtcolounname = toolStripComboBox1.SelectedItem.ToString();
                DataRow[] dr = dt.Select(dtcolounname +" like '%" + toolStripTextBox2.Text+"%'");
                DataTable dtt = new DataTable();
                dtt = dt.Clone();
                if (dr.Length > 0)
                {
                    foreach (DataRow drVal in dr)
                    {
                        dtt.ImportRow(drVal);
                    }
                }
                dataGridView1.DataSource = dtt;
                toolStripLabel3.Text = "总行数:" + (dtt.Rows.Count).ToString();
            }
             catch
            {
                    MessageBox.Show("此列暂不支持筛选", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
        }
            this.Cursor = Cursors.Default;//正常状态
        }
        //获得datatable的列名
        public static List<string> GetColumnsByDataTable(DataTable dt)
        {
            List<string> ls = new List<string>();
            if (dt.Columns.Count > 0)
            {
                int columnNum = 0;
                columnNum = dt.Columns.Count;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ls.Add(dt.Columns[i].ColumnName);
                }
            }
            return ls;
        }
        //LINQ返回Datatable
        public static DataTable ToDataTable<T>(IEnumerable<T> varlist)
        {
            DataTable dtReturn = new DataTable();
            // column names 
            PropertyInfo[] oProps = null;
            if (varlist == null)
                return dtReturn;
            foreach (T rec in varlist)
            {
                if (oProps == null)
                {
                    oProps = ((Type)rec.GetType()).GetProperties();
                    foreach (PropertyInfo pi in oProps)
                    {
                        Type colType = pi.PropertyType;
                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition()
                             == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }
                        dtReturn.Columns.Add(new DataColumn(pi.Name, colType));
                    }
                }
                DataRow dr = dtReturn.NewRow();
                foreach (PropertyInfo pi in oProps)
                {
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue
                    (rec, null);
                }
                dtReturn.Rows.Add(dr);
            }
            return dtReturn;
        }
        //导出EXCEL
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName,SqlHelper.GetDgvToTable(dataGridView1));
        }
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }
        #endregion

        #region 方法
        //private void DataGridViewToCSV()
        //{
        //    if (dataGridView1.Rows.Count == 0)
        //    {
        //        MessageBox.Show("没有数据可导出!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //    SaveFileDialog saveFileDialog = new SaveFileDialog();
        //    saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
        //    saveFileDialog.FilterIndex = 0;
        //    saveFileDialog.RestoreDirectory = true;
        //    saveFileDialog.CreatePrompt = true;
        //    saveFileDialog.FileName = null;
        //    saveFileDialog.Title = "保存";
        //    if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //    {
        //        Stream stream = saveFileDialog.OpenFile();
        //        StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.GetEncoding(-0));
        //        string strLine = "";
        //        try
        //        {
        //            //表头
        //            for (int i = 0; i < dataGridView1.ColumnCount; i++)
        //            {
        //                if (i > 0)
        //                    strLine += ",";
        //                strLine += dataGridView1.Columns[i].HeaderText;
        //            }
        //            strLine.Remove(strLine.Length - 1);
        //            sw.WriteLine(strLine);
        //            strLine = "";
        //            //表的内容
        //            for (int j = 0; j < dataGridView1.Rows.Count; j++)
        //            {
        //                strLine = "";
        //                int colCount = dataGridView1.Columns.Count;
        //                for (int k = 0; k < colCount; k++)
        //                {
        //                    if (k > 0 && k < colCount)
        //                        strLine += ",";
        //                    if (dataGridView1.Rows[j].Cells[k].Value == null)
        //                        strLine += "";
        //                    else
        //                    {
        //                        string cell = dataGridView1.Rows[j].Cells[k].Value.ToString().Trim();
        //                        //防止里面含有特殊符号
        //                        cell = cell.Replace("\"", "\"\"");
        //                        cell = "\"" + cell + "\"";
        //                        strLine += cell;
        //                    }
        //                }
        //                sw.WriteLine(strLine);
        //            }
        //            sw.Close();
        //            CSVSaveasXLS(@"C:\Users\Dell\Desktop\123.csv");
        //            stream.Close();
        //            MessageBox.Show("数据被导出到：" + saveFileDialog.FileName.ToString(), "导出完毕", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message, "导出错误", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }
        //    }
        //}
        //public static string CSVSaveasXLS(string FilePath)
        //{
        //    QuertExcel();
        //    string _NewFilePath = "";

        //    Excel.Application excelApplication;
        //    Excel.Workbooks excelWorkBooks = null;
        //    Excel.Workbook excelWorkBook = null;
        //    Excel.Worksheet excelWorkSheet = null;

        //    try
        //    {
        //        excelApplication = new Excel.ApplicationClass();
        //        excelWorkBooks = excelApplication.Workbooks;
        //        excelWorkBook = ((Excel.Workbook)excelWorkBooks.Open(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value));
        //        excelWorkSheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];
        //        excelApplication.Visible = false;
        //        excelApplication.DisplayAlerts = false;
        //        _NewFilePath = FilePath.Replace(".csv", ".xls");
        //        excelWorkBook.SaveAs(_NewFilePath, Excel.XlFileFormat.xlAddIn, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //        excelWorkBook.Close();
        //        QuertExcel();
        //        // ExcelFormatHelper.DeleteFile(FilePath);
        //        //可以不用杀掉进程QuertExcel();


        //        GC.Collect(System.GC.GetGeneration(excelWorkSheet));
        //        GC.Collect(System.GC.GetGeneration(excelWorkBook));
        //        GC.Collect(System.GC.GetGeneration(excelApplication));

        //    }
        //    catch (Exception exc)
        //    {
        //        throw new Exception(exc.Message);
        //    }

        //    finally
        //    {
        //        GC.Collect();
        //    }

        //    return _NewFilePath;
        //}

        public static bool DeleteFile(string FilePath)
        {
            try
            {
                bool IsFind = File.Exists(FilePath);
                if (IsFind)
                {
                    File.Delete(FilePath);
                }
                else
                {
                    throw new IOException("指定的文件不存在");
                }
                return true;
            }
            catch (Exception exc)
            {
                throw new Exception(exc.Message);
            }

        }
        /// <summary>
        /// 执行过程中可能会打开多个EXCEL文件 所以杀掉
        /// </summary>
        private static void QuertExcel()
        {
            Process[] excels = Process.GetProcessesByName("EXCEL");
            foreach (var item in excels)
            {
                item.Kill();
            }
        }

        //public static MemoryStream ExportDataTableToExcel(DataTable sourceTable)
        //{
        //    HSSFWorkbook workbook = new HSSFWorkbook();
        //    MemoryStream ms = new MemoryStream();
        //    int dtRowsCount = sourceTable.Rows.Count;
        //    int SheetCount = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dtRowsCount) / 65536));
        //    int SheetNum = 1;
        //    int rowIndex = 1;
        //    int tempIndex = 1; //标示
        //    ISheet sheet = workbook.CreateSheet("sheet1" + SheetNum);
        //    for (int i = 0; i < dtRowsCount; i++)
        //    {
        //        if (i == 0 || tempIndex == 1)
        //        {
        //            IRow headerRow = sheet.CreateRow(0);
        //            foreach (DataColumn column in sourceTable.Columns)
        //                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
        //        }
        //        HSSFRow dataRow = (HSSFRow)sheet.CreateRow(tempIndex);
        //        foreach (DataColumn column in sourceTable.Columns)
        //        {
        //            dataRow.CreateCell(column.Ordinal).SetCellValue(sourceTable.Rows[i][column].ToString());
        //        }
        //        if (tempIndex == 65535)
        //        {
        //            SheetNum++;
        //            sheet = workbook.CreateSheet("sheet" + SheetNum);//
        //            tempIndex = 0;
        //        }
        //        rowIndex++;
        //        tempIndex++;
        //        //AutoSizeColumns(sheet);
        //    }
        //    workbook.Write(ms);
        //    ms.Flush();
        //    ms.Position = 0;
        //    sheet = null;
        //    // headerRow = null;
        //    workbook = null;
        //    return ms;
        //}
        #endregion

        #region 分析
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            switch (NodeName)
            {
                case "产品物料查询":
                    ProductMaterialEnquiries(dt);
                    break;
                default:
                    break;
            }
        }
        //产品物料查询
        private void ProductMaterialEnquiries(DataTable dt1)
        {
            //AnalysisData.XtraForm1 pme = new AnalysisData.XtraForm1();
            //pme.Show();
        }
        #endregion

        private void toolStripTextBox2_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(toolStripTextBox2.Text))
            {
                toolStripTextBox1.Enabled = true;
            }
            else
            {
                toolStripTextBox1.Enabled = false;
            }
        }
    }
}
