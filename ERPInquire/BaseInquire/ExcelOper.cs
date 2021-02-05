using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ExcelOperation
{
    public class ExcelOperByOleDb
    {
        #region 用OleDB的方式 读取Excel 方法为静态
        public static string ConnLink = "Provider=Microsoft.Ace.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;';Data Source="; 
       
        /// <summary>
        ///  查询Excel中的所有表名
        /// </summary>
        /// <param name="ExcelFileName"></param>
        /// <returns></returns>
        public  System.Data.DataTable GetExcelAllTableName(string ExcelFileName)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            if (File.Exists(ExcelFileName))
            {
                string connStr = ConnLink + ExcelFileName;
                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();
                    dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    conn.Close();
                }
            }
            return dt;
        }

        /// <summary>
        /// 根据Excel名和表名获取数据表
        /// </summary>
        /// <param name="ExcelFileName"></param>
        /// <param name="TableName"></param>
        /// <returns></returns>
        public static System.Data.DataTable GetExcelTableColumns(string ExcelFileName, string TableName)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            if (File.Exists(ExcelFileName))
            {
                string connStr = ConnLink + ExcelFileName;
                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    //连接字符串  
                    string strCom = " SELECT * FROM [" + TableName + "$]";
                    conn.Open();
                    OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, conn);
                    myCommand.Fill(dt);
                    conn.Close();
                }
            }
            return dt;
        }
        #endregion

        //==================纠结的分割线==================//
    }
    public class ExcelOperByInterop
    {
        public System.Data.DataTable OpenExcel(string path)
        {
            DataSet ds = new DataSet();
            System.Data.DataTable dt = new System.Data.DataTable();
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            if (excel == null)
            {
                throw new Exception("无法创建Excel对象,可能您的计算机未安装Excel!");
            }
            try
            {
                Workbook wb = excel.Application.Workbooks.Open(path, missing, true, missing, missing, missing,
                    missing, missing, missing, true, missing, missing, missing, missing, missing);
                int c = wb.Worksheets.Count;
                for (int i =0; i < c; i++)
                {
                    Worksheet ws = (Worksheet) wb.Worksheets[i + 1];
                    dt = WorksheetTran(ws).Copy();
                    dt.TableName = ws.Name;
                   // ds.Tables.Add(dt);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excel.Quit();
                excel = null;
                GC.Collect();
            }
            // return ds;
            return dt;
        }

        private System.Data.DataTable WorksheetTran(Worksheet ws)
        {
            int cols = ws.UsedRange.Columns.Count;
            int rows = ws.UsedRange.Rows.Count;

            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i=0;i<cols;i++)
            {
                string str = (i + 1).ToString();
                dt.Columns.Add(str);
            }
            for (int i=0;i<rows;i++)
            {
                System.Data.DataRow dr = dt.NewRow();
                for (int j=0;j<cols;j++)
                {
                    Range tmp = (Range)ws.Cells[i + 1, j + 1];
                    dr[j] = tmp.Text;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public void  PutOutExcelByDataGridView(string title, DataGridView dgv, bool isShowExcel)
        {
            int titleColumnSpan = 0;//标题的跨列数
            string fileName = "";//保存的excel文件名
            int columnIndex = 1;//列索引
            if (dgv.Rows.Count == 0)
            {
            //    System.Windows.Forms.MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
             } 
            /*保存对话框*/
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "导出Excel(xlsx)|*.xlsx|导出Excel(xls)|*.xls";
            sfd.FileName = title + DateTime.Now.ToString("yyyyMMddhhmmss");

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                fileName = sfd.FileName;
                /*建立Excel对象*/
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                if (excel == null)
                {
                    MessageBox.Show("无法创建Excel对象,可能您的计算机未安装Excel!");
                    return;
                }
                try
                {
                    excel.Application.Workbooks.Add(true);
                    excel.Visible = isShowExcel;

                    #region 分析标题的跨列数
                    foreach (DataGridViewColumn column in dgv.Columns)
                    {
                        if (column.Visible == true)
                            titleColumnSpan++;
                    }
                    #endregion

                    #region 生成标题
                    //excel.Cells[1, 1] = title;
                    //(excel.Cells[1, 1] as Range).HorizontalAlignment = XlHAlign.xlHAlignCenter;//标题居中
                    //生成字段名称
                    columnIndex = 1;
                    for (int i = 0; i < dgv.ColumnCount; i++)
                    {
                        if (dgv.Columns[i].Visible)
                        {
                            excel.Cells[2, columnIndex] = dgv.Columns[i].HeaderText;
                            (excel.Cells[2, columnIndex] as Range).HorizontalAlignment = XlHAlign.xlHAlignCenter;//字段居中
                            columnIndex++;
                        }
                    }
                    #endregion

                    #region 填充数据
                    for (int i = 0; i < dgv.RowCount; i++)
                    {
                        columnIndex = 1;
                        for (int j = 0; j < dgv.ColumnCount; j++)
                        {
                            if (dgv.Columns[j].Visible)
                            {
                                if (dgv[j, i].ValueType == typeof(string))
                                {
                                    excel.Cells[i + 3, columnIndex] = "'" + dgv[j, i].FormattedValue;
                                }
                                else
                                {
                                    excel.Cells[i + 3, columnIndex] = dgv[j, i].Value != null ? dgv[j, i].FormattedValue : "";
                                }
                                (excel.Cells[i + 3, columnIndex] as Range).HorizontalAlignment = XlHAlign.xlHAlignLeft;//字段居中
                                columnIndex++;
                            }
                        }
                    }
                    #endregion

                    #region 新增sheet页并将数据复制

                    #endregion

                    #region 将数据转置

                    #endregion

                }
                catch { }
                finally
                {
                    excel.Quit();
                    excel = null;
                    GC.Collect();
                }
                //KillProcess("Excel");
                MessageBox.Show("导出成功!");
            }
            else
            {
            //    MessageBox.Show("导出失败");
            }
        }

        #region 杀死与Excel相关的进程
        private void KillProcess(string processName)
        {
            System.Diagnostics.Process myproc = new System.Diagnostics.Process();//得到所有打开的进程
            try
            {
                foreach (System.Diagnostics.Process thisproc in System.Diagnostics.Process.GetProcessesByName(processName))
                {
                    if (!thisproc.CloseMainWindow())
                    {
                        thisproc.Kill();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("", ex);
            }
        }
        #endregion

        #region 获取本Excel中信息
        public int GetWorksheetsCount(string path)
        {
            int res = 0;
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            if (excel == null)
            {
                throw new Exception("无法创建Excel对象,可能您的计算机未安装Excel!");
            }
            try
            {
                Workbook wb = excel.Application.Workbooks.Open(path, missing, true, missing, missing, missing,
                    missing, missing, missing, true, missing, missing, missing, missing, missing);
                res = wb.Worksheets.Count;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excel.Quit();
                excel = null;
                GC.Collect();
            }
            return res;
        }

        public int GetTarRowsCount(string path, int tarWs)
        {
            int res = 0;
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            if (excel == null)
            {
                throw new Exception("无法创建Excel对象,可能您的计算机未安装Excel!");
            }
            try
            {
                Workbook wb = excel.Application.Workbooks.Open(path, missing, true, missing, missing, missing,
                    missing, missing, missing, true, missing, missing, missing, missing, missing);
                Worksheet ws = (Worksheet)wb.Worksheets[tarWs];
                res = ws.UsedRange.Rows.Count;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excel.Quit();
                excel = null;
                GC.Collect();
            }
            return res;
        }
        #endregion

        #region 加载进度
        private Form _progressForm = new Form();
        private System.Windows.Forms.Label _lb1 = new System.Windows.Forms.Label();
        private System.Windows.Forms.Label _lb2 = new System.Windows.Forms.Label();

        public object ActiveSheet { get; private set; }

        public void ShowMessageInfo()
        {
            _lb1.Location = new System.Drawing.Point(20, 13);
            _lb2.Location = new System.Drawing.Point(20, 50);
            _progressForm.ClientSize = new System.Drawing.Size(200, 100);
            _progressForm.Controls.Add(_lb1);
            _progressForm.Controls.Add(_lb2);

            _progressForm.Text = @"加载进度";
            _progressForm.ShowIcon = false;
            _progressForm.MaximizeBox = false;
            _progressForm.MinimizeBox = false;
            _progressForm.StartPosition = FormStartPosition.CenterScreen;
            _progressForm.Show();
        }

        public void CloseMessageInfo()
        {
            _progressForm.Close();
        }
        #endregion

        //===================纠结的分隔线==================//
    }
}
