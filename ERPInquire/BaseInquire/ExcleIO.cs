using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data;
using Com.Hui.iMRP.Utils;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HPSF;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

namespace MyMIS
{
  public   class ExcleIO
    {
        public static void saveexcle(string title,DataTable dt)
      {
          #region   验证可操作性

          //定义表格内数据的行数和列数    
          int rowscount = dt.Rows.Count;
          int colscount = dt.Columns.Count;
          //行数必须大于0    
          if (rowscount <= 0)
          {
         //     System.Windows.Forms.MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
              return;
          }

          //列数必须大于0    
          if (colscount <= 0)
          {
         //     System.Windows.Forms.MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
              return;
          }

            //行数不可以大于65536    
            if (rowscount > 65536)
            {
      //          System.Windows.Forms.MessageBox.Show("数据记录数太多(最多不能超过65536条)，不能保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //列数不可以大于255    
            if (colscount > 255)
          {
       //       System.Windows.Forms.MessageBox.Show("数据记录行数太多，不能保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
              return;
          }
          #endregion

       //   SaveFileDialog saveFileDialog = new SaveFileDialog();
          SaveFileDialog sfd = new SaveFileDialog();
          sfd.Filter = "导出Excel(xlsx)|*.xlsx|导出Excel(xls)|*.xls";
          sfd.FileName = title + DateTime.Now.ToString("yyyyMMddhhmmss");
          sfd.ShowDialog();

          if (sfd.FileName.IndexOf(":") < 0) return; //被点了"取消"

          Stream myStream;
          myStream = sfd.OpenFile();
          StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
          string columnTitle = "";

          try
          {
              //写入列标题
              for (int i = 0; i < colscount; i++)
              {
                  if (i > 0)
                  {
                      columnTitle += "\t";
                  }
                  columnTitle += dt.Columns[i].ColumnName;
              }
              sw.WriteLine(columnTitle);

              //写入列内容
              for (int j = 0; j < rowscount; j++)
              {
                  string columnValue = "";
                  for (int k = 0; k < colscount; k++)
                  {
                      if (k > 0)
                      {
                          columnValue += "\t";
                      }
                      if (dt.Rows[j][k] == null)
                          columnValue += "";
                      else
                      {
                          if (dt.Rows[j][k].GetType() == typeof(string) && dt.Rows[j][k].ToString().StartsWith("0"))
                          {
                              columnValue += "'" + dt.Rows[j][k].ToString();
                          }
                          else
                              columnValue += dt.Rows[j][k].ToString();
                      }
                  }
                  sw.WriteLine(columnValue);
              }
              sw.Close();
              myStream.Close();
                MessageBox.Show("导出成功1");
          }
          catch (Exception ex)
          {
              System.Windows.Forms.MessageBox.Show(ex.ToString());
          }
          finally
          {
              sw.Close();
              myStream.Close();
          }
      }
        /***************/
        public static void OutExcel(string title,DataTable dt)
        {
            //定义表格内数据的行数和列数    
            int rowscount = dt.Rows.Count;
            int colscount = dt.Columns.Count;
            //行数必须大于0    
            if (rowscount <= 0)
            {
           //     System.Windows.Forms.MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //列数必须大于0    
            if (colscount <= 0)
            {
           //     System.Windows.Forms.MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //列数不可以大于255    
            if (colscount > 255)
            {
           // System.Windows.Forms.MessageBox.Show("数据记录行数太多，不能保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            //sfd.Filter = "导出Excel(xlsx)|*.xlsx|导出Excel(xls)|*.xls";
            sfd.Filter = "导出Excel(xls)|*.xls";
            sfd.FileName = title + DateTime.Now.ToString("yyyyMMddhhmmss");
            sfd.ShowDialog();

            if (sfd.FileName.IndexOf(":") < 0) return; //被点了"取消"

            try
            {
                using (StreamWriter sw = new StreamWriter(sfd.FileName, false, Encoding.GetEncoding("gb2312")))
                {
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        sb.Append(dt.Columns[i].ColumnName.ToString() + "\t");
                    }
                    sb.Append(Environment.NewLine);

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        System.Windows.Forms.Application.DoEvents();

                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            sb.Append(dt.Rows[r][c].ToString() + "\t");
                        }
                        sb.Append(Environment.NewLine);
                    }
                    sw.Write(sb.ToString());
                    sw.Flush();
                    sw.Close();
                    MessageBox.Show("导出成功，总共导出" + dt.Rows.Count + "条数据!");
                }
            }
            catch
            {
                MessageBox.Show("导出失败!");
            }
        }
        /*******************/
        public static void DataTableToExcel(string name, DataTable dt)
        {
            Random rnd = new Random();
            Excel.ExcelHelper myhelper = new Excel.ExcelHelper();
   //         byte[] data = myhelper.DataTable2Excel(dt, name);
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "导出Excel(xls)|*.xls";
            sfd.FileName = name + DateTime.Now.ToString("yyyyMMddhhmmss");
            sfd.ShowDialog();
            if (!File.Exists(sfd.FileName))
            {
                //FileStream fs = new FileStream(sfd.FileName, FileMode.CreateNew);
                //fs.Write(data, 0, data.Length);
                //fs.Close();
            }
        }

        //public static void TableToExcelForXLSX2007(string name,DataTable dt)
        //{
        //    SaveFileDialog sfd = new SaveFileDialog();
        //    sfd.Filter = "导出Excel(xls)|*.xls";
        //    sfd.FileName = name + DateTime.Now.ToString("yyyyMMddhhmmss");
        //    sfd.ShowDialog();
        //    if (sfd.FileName.IndexOf(":") < 0) return; //被点了"取消"

        //    XSSFWorkbook xssfworkbook = new XSSFWorkbook();//建立Excel2007对象
        //    ISheet sheet = xssfworkbook.CreateSheet(name);//新建一个名称为sheetname的工作簿
        //    //设置列名
        //    IRow row = sheet.CreateRow(0);
        //    for (int i = 0; i < dt.Columns.Count; i++)
        //    {
        //        ICell cell = row.CreateCell(i);
        //        cell.SetCellValue(dt.Columns[i].ColumnName);
        //    }
        //    //单元格赋值
        //    for (int i = 0; i < dt.Rows.Count; i++)
        //    {
        //        IRow row1 = sheet.CreateRow(i + 1);
        //        for (int j = 0; j < dt.Columns.Count; j++)
        //        {
        //            ICell cell = row1.CreateCell(j);
        //            cell.SetCellValue(dt.Rows[i][j].ToString());
        //        }
        //    }
        //    using (System.IO.Stream stream = File.OpenWrite(sfd.FileName))
        //    {
        //        //写入文件
        //        try
        //        {
        //            xssfworkbook.Write(stream);
        //            stream.Close();
        //            MessageBox.Show("导出成功！");
        //        }
        //        catch
        //        {
        //            MessageBox.Show("导出失败！");
        //        }
        //    }
        //}

        //public static void TableToExcelForXLSX2003(string name,DataTable dt)
        //{
        //    SaveFileDialog sfd = new SaveFileDialog();
        //    sfd.Filter = "导出Excel(xls)|*.xls";
        //    sfd.FileName = name + DateTime.Now.ToString("yyyyMMddhhmmss");
        //    sfd.ShowDialog();
        //    if (sfd.FileName.IndexOf(":") < 0) return; //被点了"取消"
        //    HSSFWorkbook xssfworkbook = new HSSFWorkbook();//建立Excel2003对象
        //    HSSFSheet sheet = (HSSFSheet)xssfworkbook.CreateSheet(name);//新建一个名称为sheetname的工作簿
        //    //设置列名
        //    HSSFRow row = (HSSFRow)sheet.CreateRow(0);
        //    for (int i = 0; i < dt.Columns.Count; i++)
        //    {
        //        ICell cell = (ICell)row.CreateCell(i);
        //        cell.SetCellValue(dt.Columns[i].ColumnName);
        //    }
        //    //单元格赋值
        //    for (int i = 0; i < dt.Rows.Count; i++)
        //    {
        //        IRow row1 = sheet.CreateRow(i + 1);
        //        for (int j = 0; j < dt.Columns.Count; j++)
        //        {
        //            ICell cell = row1.CreateCell(j);
        //            cell.SetCellValue(dt.Rows[i][j].ToString());
        //        }
        //    }
        //    using (System.IO.Stream stream = File.OpenWrite(sfd.FileName))
        //    {
        //        try
        //        {
        //            xssfworkbook.Write(stream);
        //            stream.Close();
        //            MessageBox.Show("导出成功！");
        //        }
        //        catch
        //        {
        //            MessageBox.Show("导出失败！");
        //        }
        //    }
        //}

        public static void Export2Excel(string nodename, DataTable dt)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = nodename + DateTime.Now.ToString("yyyyMMddhhmmss");
            dlg.Filter = "xlsx files(*.xlsx)|*.xlsx|xls files(*.xls)|*.xls|All files(*.*)|*.*";
            dlg.ShowDialog();
            if (dlg.FileName.IndexOf(":") < 0) return; //被点了"取消"
            ExcelHelper helper = new ExcelHelper(dlg.FileName);
            //int i=helper.DataTableToExcel(dt, "sheet1", true);
            DataTableToExcel1(dt, dlg.FileName, "sheet1", true); ;
            MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public static void Export2Excel1(string nodename, DataTable dt)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = nodename + DateTime.Now.ToString("yyyyMMddhhmmss");
            dlg.Filter = "xlsx files(*.xlsx)|*.xlsx|xls files(*.xls)|*.xls|All files(*.*)|*.*";
            dlg.ShowDialog();
            if (dlg.FileName.IndexOf(":") < 0) return; //被点了"取消"
            ExcelHelper helper = new ExcelHelper(dlg.FileName);
            //int i=helper.DataTableToExcel(dt, "sheet1", true);
            DataTableToExcel2(dt, dlg.FileName, "sheet1", true);
            MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public static int DataTableToExcel(DataTable data,string fileName, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            IWorkbook workbook = null;
            FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else
                workbook = new HSSFWorkbook();
            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        public static int DataTableToExcel1(DataTable dt, string fileName, string sheetName, bool isColumnWritten)
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

        public static int DataTableToExcel2(DataTable dt, string fileName, string sheetName, bool isColumnWritten)
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
                //设置导出文件的背景色
                ICellStyle s = workbook.CreateCellStyle();
                s.FillForegroundColor = HSSFColor.Red.Index;
                s.FillPattern = FillPattern.SolidForeground;
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    IRow row = sheet.CreateRow(num);
                    //判断是否产品档案进行审核
                    try
                    {
                        if (CheckCPSH(dt.Rows[j]["产品编码"].ToString()).Equals(true))
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if ((dt.Columns[i].ColumnName.Contains("/") || dt.Columns[i].ColumnName.Contains("-") || dt.Columns[i].ColumnName.Contains("量") || dt.Columns[i].ColumnName.Contains("存") || dt.Columns[i].ColumnName.Contains("印刷")) && dt.Rows[j][i] != DBNull.Value)
                                {
                                    try
                                    {
                                        row.CreateCell(i).SetCellValue(Convert.ToDouble(dt.Rows[j][i]));
                                        row.GetCell(i).CellStyle = s;

                                    }
                                    catch
                                    {
                                        row.CreateCell(i).SetCellValue(dt.Rows[j][i].ToString());
                                        row.GetCell(i).CellStyle = s;
                                    }
                                }
                                else
                                {
                                    row.CreateCell(i).SetCellValue(dt.Rows[j][i].ToString());
                                    row.GetCell(i).CellStyle = s;
                                }
                            }
                        else
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if ((dt.Columns[i].ColumnName.Contains("/") || dt.Columns[i].ColumnName.Contains("-") || dt.Columns[i].ColumnName.Contains("量") || dt.Columns[i].ColumnName.Contains("存") || dt.Columns[i].ColumnName.Contains("印刷")) && dt.Rows[j][i] != DBNull.Value)
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
                        }
                    }
                    catch
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            if ((dt.Columns[i].ColumnName.Contains("/") || dt.Columns[i].ColumnName.Contains("-") || dt.Columns[i].ColumnName.Contains("量") || dt.Columns[i].ColumnName.Contains("存") || dt.Columns[i].ColumnName.Contains("印刷")) && dt.Rows[j][i] != DBNull.Value)
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

        //判断产品档案是否被审核
        private static bool CheckCPSH(string s)
        {
            string sql = "SELECT SHF FROM JCZL_CPBOM_M WHERE CPBH='" + s + "'";
            DataTable dcf=SqlHelper.ExecuteDataTable(sql);
            string ss = dcf.Rows[0][0].ToString();
            if (ss.Equals("False"))
            {
                //未审核
                return true;
            }
            else
            {
                //已审核
                return false;
            }
        }
    }
}
