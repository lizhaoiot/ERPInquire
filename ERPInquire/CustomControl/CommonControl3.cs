using System;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ERPInquire.CustomControl
{
    public partial class CommonControl3 : UserControl
    {
        #region 变量

        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dtNew = new DataTable();
        #endregion

        #region  定量

        string path1 = "Template";
        public string[] names;
        public string[] attributes;
        public string nodename;
        private string fileAddress;
        private string tableTempName = null;
        public string ProName;
        public string NodeName;
        #endregion

        #region 提示信息

        private static string str1 = "文件下载失败";

        #endregion

        #region 初始化
        public CommonControl3()
        {
            InitializeComponent();
            toolStripButton2.Enabled = false;
            toolStripTextBox1.Enabled = false;
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行被选中
        }

        #endregion

        #region 窗体事件

        //从本地复制文件到桌面
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
          FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fd = dlg.SelectedPath;
                CopyDir(path1, fd);
            }
        }

        //上传查询数据
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            toolStripButton2.Enabled = true;
            toolStripTextBox1.Enabled = true;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();  //显示选择文件对话框  
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                toolStripTextBox1.Text = openFileDialog1.FileName;
                fileAddress= openFileDialog1.FileName;
            }
        }

        //查询数据
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
           // ExcelOperation.ExcelOperByInterop excel = new ExcelOperation.ExcelOperByInterop();
            DataTable dt = new DataTable();
            try
            {
                dt = ImportExcel(fileAddress);
            }
            catch
            {
                MessageBox.Show("文件打开失败，请核对文件上传路径！");
                return;
            }
            CreateTempTable();
            try
            {
                SqlHelper.BulkToDB(Convertdt(dt), tableTempName, names);
                dt2 = CommandPro();
                dataGridView1.DataSource = dt2;
                toolStripLabel4.Text ="总行数:"+ Convert.ToString(dt2.Rows.Count);
                dataGridView1.RowsDefaultCellStyle.BackColor = Color.Aquamarine;
                dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Bisque;
                dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
                dataGridView1.Refresh();
                DeleteTempTable();
            }
            catch(Exception ex)
            {
                DeleteTempTable();
                MessageBox.Show("查询失败");
            }
            toolStripButton2.Enabled = false;
        }
        private DataTable Convertdt(DataTable dt)
        {
            DataColumn dc = null;
            dc = dtNew.Columns.Add("产品编号", Type.GetType("System.String"));
            dc = dtNew.Columns.Add("产品数量", Type.GetType("System.Double"));
            dtNew = dt.Copy();
            return dtNew;
        }
        //导出EXCEL
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
             MyMIS.ExcleIO.Export2Excel(NodeName, dt2);
        }
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
        //    MyMIS.ExcleIO.TableToExcelForXLSX2003(NodeName, dt2);
        }
        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (toolStripTextBox1.Text != "")
            {
                toolStripButton2.Enabled = true;
            }
            else
            {
                toolStripButton2.Enabled = false;
            }
        }
        #endregion

        #region 方法
        //复制文件
        private void CopyDir(string srcPath, string aimPath)
        {
            try
            {
                if (aimPath[aimPath.Length - 1] != System.IO.Path.DirectorySeparatorChar)
                {
                    aimPath += System.IO.Path.DirectorySeparatorChar;
                }
                if (!System.IO.Directory.Exists(aimPath))
                {
                    System.IO.Directory.CreateDirectory(aimPath);
                }
                string[] fileList = System.IO.Directory.GetFileSystemEntries(srcPath);
                // 遍历所有的文件和目录
                foreach (string file in fileList)
                {
                      string a1 = "Template\\" + nodename + ".xlsx";
                      if (file.Equals(a1))
                     { 
                       if (System.IO.Directory.Exists(file))
                       {
                           CopyDir(file, aimPath + System.IO.Path.GetFileName(file));
                       }
                       else
                       {
                           System.IO.File.Copy(file, aimPath + System.IO.Path.GetFileName(file), true);
                       }
                     }
                }
                System.Diagnostics.Process.Start(aimPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(str1);
            }
        }

        /// <summary>
        /// 删除指定后缀名的文件
        /// </summary>
        /// <param name="directory">删除的绝对路径</param>
        /// <param name="masks">后缀名的数组</param>
        /// <param name="searchSubdirectories">是否需要递归删除</param>
        /// <param name="ignoreHidden">是否忽略隐藏文件</param>
        /// <param name="deletedFileCount">总共删除文件数</param>
        public void DeleteFiles(string directory, string[] masks, bool searchSubdirectories, bool ignoreHidden, ref int deletedFileCount)
        {
            //先删除当前目录下指定后缀名的所有文件
            foreach (string file in Directory.GetFiles(directory, "*.*"))
            {
                if (!(ignoreHidden && (File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden))
                {
                    foreach (string mask in masks)
                    {
                        if (Path.GetExtension(file) == mask)
                        {
                            File.Delete(file);
                            deletedFileCount++;
                        }
                    }
                }
            }
            //如果需要对子目录进行处理，则对子目录也进行递归操作
            if (searchSubdirectories)
            {
                string[] childDirectories = Directory.GetDirectories(directory);
                foreach (string dir in childDirectories)
                {
                    if (!(ignoreHidden && (File.GetAttributes(dir) & FileAttributes.Hidden) == FileAttributes.Hidden))
                    {
                        DeleteFiles(dir, masks, searchSubdirectories, ignoreHidden, ref deletedFileCount);
                    }
                }
            }
        }
        #endregion

        #region 临时表
        //建立临时表
        private void CreateTempTable()
        {
            Random rd = new Random();
            string rds = Convert.ToString(rd.Next(1, 1000000));
            SqlHelper.GetConnection();
            string sql = "create table table" + rds +" (";
            for (int i = 0; i < names.Length; i++)
            {
                sql = sql + names[i] +" "+ attributes[i]+ ",";
            }
            sql = sql.Substring(0, sql.Length - 1);
            sql = sql + ")";
            try
            {
                SqlHelper.ExecCommand(sql);
                tableTempName = "table" + rds;
            }
            catch
            {
                MessageBox.Show("现在查询人数过多，请稍后查询");
            }
             SqlHelper.GetConnection().Close();
        }
        //删除临时表
        private void DeleteTempTable()
        {
            SqlHelper.GetConnection();
            string sql = "drop table " + tableTempName;
            try
            {
                SqlHelper.ExecCommand(sql);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            tableTempName = null;
            SqlHelper.GetConnection().Close();
        }
        private DataTable CommandPro()
        {
            DataTable dt1 = new DataTable();
            string connectionString = "data source=192.168.0.97; Database=hy;user id=sa; password=";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = ProName;
                        SqlParameter[] para ={
                    new SqlParameter("@DAH",SqlDbType.VarChar),
                };
                        para[0].Value = tableTempName;
                        cmd.CommandTimeout = 60 * 60 * 100000;
                        cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dt1);
                        conn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }         
             }
            return dt1;
         }
        #endregion

        #region EXCEL
        public static DataTable ImportExcel(string filePath)
        {
            DataTable dt = new DataTable();
            using (FileStream fsRead = System.IO.File.OpenRead(filePath))
            {
                IWorkbook wk = null;
                //获取后缀名
                string extension = filePath.Substring(filePath.LastIndexOf(".")).ToString().ToLower();
                //判断是否是excel文件
                if (extension == ".xlsx" || extension == ".xls")
                {
                    //判断excel的版本
                    if (extension == ".xlsx")
                    {
                        wk = new XSSFWorkbook(fsRead);
                    }
                    else
                    {
                        wk = new HSSFWorkbook(fsRead);
                    }

                    //获取第一个sheet
                    ISheet sheet = wk.GetSheetAt(0);
                    //获取第一行
                    IRow headrow = sheet.GetRow(0);
                    //创建列
                    for (int i = headrow.FirstCellNum; i < headrow.Cells.Count; i++)
                    {
                        //  DataColumn datacolum = new DataColumn(headrow.GetCell(i).StringCellValue);
                        DataColumn datacolum = new DataColumn("F" + (i + 1));
                        dt.Columns.Add(datacolum);
                    }
                    //读取每行,从第二行起
                    for (int r = 1; r <= sheet.LastRowNum; r++)
                    {
                        bool result = false;
                        DataRow dr = dt.NewRow();
                        //获取当前行
                        IRow row = sheet.GetRow(r);
                        //读取每列
                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            ICell cell = row.GetCell(j); //一个单元格
                            dr[j] = GetCellValue(cell); //获取单元格的值
                                                        //全为空则不取
                            if (dr[j].ToString() != "")
                            {
                                result = true;
                            }
                        }
                        if (result == true)
                        {
                            dt.Rows.Add(dr); //把每行追加到DataTable
                        }
                    }
                }
            }
            return dt;
        }
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 这里类型注意一下，不同版本NPOI大小写可能不一样,有的版本是Blank（首字母大写)
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (HSSFDateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型
                    return cell.ToString();//
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }
        #endregion

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}  