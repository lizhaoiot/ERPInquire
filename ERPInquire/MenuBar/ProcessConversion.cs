using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;


namespace ERPInquire.MenuBar
{
    public partial class ProcessConversion : Form
    {
        #region 定量

        #endregion

        #region 变量
        //精简dt
        private DataTable dtt1 = new DataTable();
        private DataTable dtt2 = new DataTable();
        //主计划输出产品别工序正数  
        private DataTable dtp1 = new DataTable();
        //主计划输出工序别正数      
        private DataTable dtp2 = new DataTable();
        //主计划输出产品别物料需求 
        private DataTable dtw11 = new DataTable();
        //主计划输出物料别需求 
        private DataTable dtw21 = new DataTable();
        private static DataTable dtp71 = new DataTable();
        private static DataTable dtp72 = new DataTable();
        private static DataTable dtp73 = new DataTable();
        private static DataTable dtp74 = new DataTable();

        private static DataTable ndt1 = new DataTable();
        private static DataTable ndt2 = new DataTable();
        private static DataTable ndt3 = new DataTable();
        private static DataTable ndt4 = new DataTable();

        private string fileAddress;
        private static ProcessConversion frm = null;
        string[] strColumns = null;

        private int ox;
        private int oy;

        //分隔表格标志
        private int index1 = 0;
        private int index2 = 0;
        #endregion

        #region 
        DataColumn d10 = new DataColumn("产品编码");
        DataColumn d11 = new DataColumn("产品名称");
        DataColumn d12 = new DataColumn("CPBH");
        DataColumn d13 = new DataColumn("CPMC");
        #endregion

        #region 初始化
        private ProcessConversion()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            InitializeComponent();
        }
        public static ProcessConversion CreateInstrance()
        {
            if (frm == null || frm.IsDisposed)
            {
                frm = new ProcessConversion();
            }
            return frm;
        }
        private void ProcessConversion_Load(object sender, EventArgs e)
        {
            //button2.Enabled = false;
            button3.Enabled = false;
            label2.Text = ReadDateTime();
        }
        #endregion

        #region 窗体事件

        //更新本地数据库
        private void button1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            try
            {
                SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
                string s = connect.GetConnectStr("ERP");
                string connectionString = s;

                #region 插入工序基础资料
                InsertGXData(null, false);
                #endregion

                #region 插入物料基础资料
                InsertWLData(null, false);
                #endregion

                WriteDateTime();
                label2.Text = ReadDateTime();
                MessageBox.Show("本地数据库更新成功！更新工序基础资料表：" + ox + "条数据，更新物料基础资料表：" + oy + "条数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                //button2.Enabled = true;
                //button3.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("本地数据库更新失败！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //MessageBox.Show(ex.Message);
            }
            this.Cursor = Cursors.Default;//正常       
        }
        //获取列名
        public static string[] GetColumnsByDataTable(DataTable dt)
        {
            string[] strColumns = null;
            if (dt.Columns.Count > 0)
            {
                int columnNum = 0;
                columnNum = dt.Columns.Count - 13;
                strColumns = new string[columnNum];
                for (int i = 0; i < dt.Columns.Count - 13; i++)
                {
                    strColumns[i] = dt.Columns[i + 13].ColumnName;
                }
            }
            return strColumns;
        }
        //导入文件
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();  //显示选择文件对话框  
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.Cursor = Cursors.WaitCursor;//等待
                fileAddress = openFileDialog1.FileName;
            }
            DataTable dt = new DataTable();
            try
            {
                if (openFileDialog1.FileName.IndexOf(":") < 0) return; //被点了"取消"

                dt = ImportExcel(fileAddress);
                dtt1 = StreamlineDT1(dt);
                dtt2 = StreamlineDT2(dt);
                strColumns = GetColumnsByDataTable(dtt1);

                #region 调取本地数据库数据表           
                //string sql1 = "select * from BasicProcessInformation";
                //DataTable dtGX = Tools.Common.SQLiteHelper.ExecuteDatatable(sql1);
                //string sql2 = "select * from BasicMaterials";
                //DataTable dtWL = Tools.Common.SQLiteHelper.ExecuteDatatable(sql2);
                #endregion

                #region old
                ////计算产品别工序正数
                //dtp1 = CalcProductProcessNumber(dtt1, dtGX);
                //dtp71 = dtp1.Copy();
                ////计算工序别正数
                //dtp2 = CalcProcessNumber(dtp1);
                //dtp72 = dtp2.Copy();
                ////计算产品别物料需求
                //dtw11 = CalcProductMaterialNumber(dtt2, dtWL);
                //dtp73 = dtw11.Copy();
                ////计算物料别
                //dtw21 = CalcMaterialNumber(dtw11);
                //dtp74 = dtw21.Copy();

                #endregion

                #region new 
                //插入工序临时表
                nCalcProductProcessNumber(dtt1);
                //插入物料临时表
                nCalcMaterialNumber(dtt2);

                //计算产品别工序正数
                ndt1 = n1();
                //回绑列名
                for (int i = 0; i < strColumns.Length; i++)
                {
                    ndt1.Columns[i + 16].ColumnName = strColumns[i];
                }
                //计算工序别正数
                ndt2 = n2();
                //回绑列名
                for (int i = 0; i < strColumns.Length; i++)
                {
                    ndt2.Columns[i + 3].ColumnName = strColumns[i];
                }
                //计算产品别物料需求
                ndt3 = n3();
                //回绑列名
                for (int i = 0; i < strColumns.Length; i++)
                {
                    ndt3.Columns[i + 18].ColumnName = strColumns[i];
                }
                //计算物料别
                ndt4 = n4();
                //回绑列名
                for (int i = 0; i < strColumns.Length; i++)
                {
                    ndt4.Columns[i + 6].ColumnName = strColumns[i];
                }
                #endregion

                this.Cursor = Cursors.Default;
                button3.Enabled = true;
                MessageBox.Show("导入文件成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                ProcessConversionPreview pp = new ProcessConversionPreview();
                ProcessConversionPreview.dt1 = n5();
                ProcessConversionPreview.dt2 = ndt2;
                ProcessConversionPreview.dt3 = n6();
                ProcessConversionPreview.dt4 = ndt4;
                pp.Show();
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
         //去除全为空的列
         private DataTable DeleteNullColoumn(DataTable dt)
        {
            List<DataRow> removelist = new List<DataRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {

                        rowdataisnull = false;
                    }
                }
                if (rowdataisnull)
                {
                    removelist.Add(dt.Rows[i]);
                }

            }
            for (int i = 0; i < removelist.Count; i++)
            {
                dt.Rows.Remove(removelist[i]);
            }
            return dt;
        }
         //导出文件
         private void button3_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel1("主计划输出产品别工序正数明细",ndt1);
            MyMIS.ExcleIO.Export2Excel1("主计划输出汇总工序别正数", ndt2);
            MyMIS.ExcleIO.Export2Excel1("主计划输出产品物料别明细",cgg1(ndt3));
            MyMIS.ExcleIO.Export2Excel1("主计划输出汇总物料别", cgg2(ndt4));
            button3.Enabled = false;
        }
        #endregion


        #region 拆分合并表格
        private DataTable cgg1(DataTable d1)
        {
            DataTable dcc1 = new DataTable();
            DataSet dss1 = SplitDataTable(d1, d1.Rows.Count/2);
            DataColumn d11 = new DataColumn("序号");
            DataColumn d12 = new DataColumn("产品编码");
            DataColumn d13 = new DataColumn("物料编码");
            DataColumn d14 = new DataColumn("客户名称");
            DataColumn d15 = new DataColumn("物料单位");
            DataColumn[] dc1 = new DataColumn[] { d11,d12,d13,d14,d15 };
            DataColumn[] dc2 = new DataColumn[] { d11, d12, d13, d14, d15 };
            dcc1 = Join(dss1.Tables[0], dss1.Tables[1],dc1,dc2);
            return dcc1;
        }
        private DataTable cgg2(DataTable d1)
        {
            DataTable dcc1 = new DataTable();
            DataSet dss1 = SplitDataTable(d1, d1.Rows.Count/2);
            DataColumn d11 = new DataColumn("物料大类");
            DataColumn d12 = new DataColumn("物料子类");
            DataColumn d13 = new DataColumn("物料编码");
            DataColumn d14 = new DataColumn("物料名称");
            DataColumn d15 = new DataColumn("物料单位");
            DataColumn[] dc1 = new DataColumn[] { d11,d12,d13,d14,d15};
            DataColumn[] dc2 = new DataColumn[] { d11, d12, d13, d14, d15 };
            dcc1 = Join(dss1.Tables[0], dss1.Tables[1], dc1, dc2);
            return dcc1;
        }
        #endregion

        #region 方法

        //产品别工序正数
        private DataTable CalcProductProcessNumber(DataTable dt, DataTable dtGX)
        {
            DataTable d1 = new DataTable();
            DataColumn[] dc1 = new DataColumn[] { d10, d11 };
            DataColumn[] dc2 = new DataColumn[] { d12, d13 };
            d1 = JoinTwoTable(dt, dtGX, dc1, dc2);
            //对d1的行数顺序进行调整,去除不要行
            d1.Columns.Remove("CPBH");
            d1.Columns.Remove("CPMC");
            d1.Columns.Remove("ZBXH_BJBH");
            d1.Columns.Remove("BJMC");
            d1.Columns.Remove("SJC");
            d1.Columns.Remove("SJK");
            d1.Columns.Remove("ZMYS");
            d1.Columns.Remove("FMYS");
            d1.Columns.Remove("FSMC");

            d1.Columns["DL"].ColumnName = "物料得率";
            d1.Columns["GXBH"].ColumnName = "工序编码";
            d1.Columns["GBLB"].ColumnName = "工序类别";
            d1.Columns["GXMC"].ColumnName = "工序名称";
            d1.Columns["BHCS"].ColumnName = "变化系数";

            d1.Columns["序号"].SetOrdinal(0);
            d1.Columns["客户名称"].SetOrdinal(1);
            d1.Columns["产品编码"].SetOrdinal(2);
            d1.Columns["产品名称"].SetOrdinal(3);
            d1.Columns["订单余量合计"].SetOrdinal(4);
            d1.Columns["送货数量合计"].SetOrdinal(5);
            d1.Columns["库存"].SetOrdinal(6);
            d1.Columns["已印刷"].SetOrdinal(7);
            d1.Columns["待印刷"].SetOrdinal(8);
            d1.Columns["往期未转化为施工单的订单"].SetOrdinal(9);
            d1.Columns["判断是否超产需要扣数"].SetOrdinal(10);
            d1.Columns["区分1"].SetOrdinal(11);
            d1.Columns["交货天数"].SetOrdinal(12);
            d1.Columns["物料得率"].SetOrdinal(13);
            d1.Columns["工序编码"].SetOrdinal(14);
            d1.Columns["工序类别"].SetOrdinal(15);
            d1.Columns["工序名称"].SetOrdinal(16);
            d1.Columns["变化系数"].SetOrdinal(17);
            for (int i = 0; i < d1.Rows.Count; i++)
            {
                for (int j = 18; j < d1.Columns.Count; j++)
                {
                    d1.Rows[i][j] = Convert.ToString(Convert.ToInt64(Convert.ToDecimal(d1.Rows[i][j]) * Convert.ToDecimal(d1.Rows[i][13]) * Convert.ToDecimal(d1.Rows[i][17])));
                }
            }
            d1.Columns.Remove("物料得率");
            d1.Columns.Remove("变化系数");
            return d1;
        }
        //工序别正数
        private DataTable CalcProcessNumber(DataTable dt)
        {
            dt.Columns.Remove("序号");
            dt.Columns.Remove("客户名称");
            dt.Columns.Remove("产品编码");
            dt.Columns.Remove("产品名称");
            dt.Columns.Remove("订单余量合计");
            dt.Columns.Remove("送货数量合计");
            dt.Columns.Remove("库存");
            dt.Columns.Remove("已印刷");
            dt.Columns.Remove("待印刷");
            dt.Columns.Remove("往期未转化为施工单的订单");
            dt.Columns.Remove("判断是否超产需要扣数");
            dt.Columns.Remove("区分1");
            dt.Columns.Remove("交货天数");
            DataTable d2 = new DataTable();
            string s1 = string.Empty;
            string s2 = string.Empty;
            string s3 = string.Empty;
            string s4 = string.Empty;
            string s5 = string.Empty;
            string s6 = string.Empty;
            int index = dt.Columns.Count - 3;
            long[] str = new long[index];
            long[] str1 = new long[index];
            d2 = dt.Clone();
            d2.Clear();
            int ijk = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                s1 = dt.Rows[i][0].ToString();
                s2 = dt.Rows[i][1].ToString();
                s3 = dt.Rows[i][2].ToString();
                for (int j = i + 1; j < dt.Rows.Count; j++)
                {
                    s4 = dt.Rows[j][0].ToString();
                    s5 = dt.Rows[j][1].ToString();
                    s6 = dt.Rows[j][2].ToString();
                    if (s1 == s4 && s2 == s5 && s3 == s6)
                    {
                        for (int k = 0; k < index; k++)
                        {
                            try
                            {
                                str1[k] = Convert.ToInt64(dt.Rows[i][k + 3]);
                                str[k] = Convert.ToInt64(dt.Rows[j][k + 3]) + str[k];
                            }
                            catch
                            {
                                continue;
                            }
                        }
                        for (int p = 0; p < dt.Columns.Count; p++)
                        {
                            dt.Rows[j][p] = "";
                        }
                    }
                    else if (s1 != s4 && s2 != s5 && s3 != s6)
                    {
                        for (int k = 0; k < index; k++)
                        {
                            try
                            {
                                str1[k] = Convert.ToInt64(dt.Rows[i][k + 3]);
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
                d2.Rows.Add();
                for (int l = 0; l < index; l++)
                {
                    d2.Rows[ijk][0] = s1;
                    d2.Rows[ijk][1] = s2;
                    d2.Rows[ijk][2] = s3;
                    d2.Rows[ijk][l + 3] = str1[l] + str[l];
                }
                for (int k = 0; k < index; k++)
                {
                    str1[k] = 0;
                    str[k] = 0;
                }
                ijk = ijk + 1;
            }
            DataRow[] foundRow;
            foundRow = d2.Select("工序编码=''");
            foreach (DataRow row in foundRow)
            {
                d2.Rows.Remove(row);
            }
            return d2;
        }
        //产品别物料
        private DataTable CalcProductMaterialNumber(DataTable dt, DataTable dtWL)
        {
            DataTable d3 = new DataTable();
            DataColumn[] dc1 = new DataColumn[] { d10, d11 };
            DataColumn[] dc2 = new DataColumn[] { d12, d13 };
            d3 = JoinTwoTable(dt, dtWL, dc1, dc2);

            d3.Columns.Remove("CPBH");
            d3.Columns.Remove("CPMC");
            d3.Columns.Remove("BJMC");
            d3.Columns.Remove("SJC");
            d3.Columns.Remove("SJK");
            d3.Columns.Remove("BZYL");

            d3.Columns["DL"].ColumnName = "物料得率";
            d3.Columns["WLDL"].ColumnName = "物料大类";
            d3.Columns["WLZL"].ColumnName = "物料子类";
            d3.Columns["WLBH"].ColumnName = "物料编号";
            d3.Columns["WLMC"].ColumnName = "物料名称";

            d3.Columns["序号"].SetOrdinal(0);
            d3.Columns["客户名称"].SetOrdinal(1);
            d3.Columns["产品编码"].SetOrdinal(2);
            d3.Columns["产品名称"].SetOrdinal(3);
            d3.Columns["订单余量合计"].SetOrdinal(4);
            d3.Columns["送货数量合计"].SetOrdinal(5);
            d3.Columns["库存"].SetOrdinal(6);
            d3.Columns["已印刷"].SetOrdinal(7);
            d3.Columns["待印刷"].SetOrdinal(8);
            d3.Columns["往期未转化为施工单的订单"].SetOrdinal(9);
            d3.Columns["判断是否超产需要扣数"].SetOrdinal(10);
            d3.Columns["区分1"].SetOrdinal(11);
            d3.Columns["交货天数"].SetOrdinal(12);
            d3.Columns["物料大类"].SetOrdinal(13);
            d3.Columns["物料子类"].SetOrdinal(14);
            d3.Columns["物料编号"].SetOrdinal(15);
            d3.Columns["物料名称"].SetOrdinal(16);
            d3.Columns["物料得率"].SetOrdinal(17);

            for (int i = 0; i < d3.Rows.Count; i++)
            {
                for (int j = 18; j < d3.Columns.Count; j++)
                {
                    d3.Rows[i][j] = Convert.ToString(Convert.ToInt64(Convert.ToDecimal(d3.Rows[i][j]) * Convert.ToDecimal(d3.Rows[i][17])));
                }
            }
            d3.Columns.Remove("物料得率");
            int ij = d3.Rows.Count;
            return d3;
        }
        //物料别
        private DataTable CalcMaterialNumber(DataTable dt)
        {
            DataTable d4 = new DataTable();
            dt.Columns.Remove("序号");
            dt.Columns.Remove("客户名称");
            dt.Columns.Remove("产品编码");
            dt.Columns.Remove("产品名称");
            dt.Columns.Remove("订单余量合计");
            dt.Columns.Remove("送货数量合计");
            dt.Columns.Remove("库存");
            dt.Columns.Remove("已印刷");
            dt.Columns.Remove("待印刷");
            dt.Columns.Remove("往期未转化为施工单的订单");
            dt.Columns.Remove("判断是否超产需要扣数");
            dt.Columns.Remove("交货天数");

            string s1 = string.Empty;
            string s2 = string.Empty;
            string s3 = string.Empty;
            string s31 = string.Empty;
            string s32 = string.Empty;

            string s4 = string.Empty;
            string s5 = string.Empty;
            string s6 = string.Empty;
            string s61 = string.Empty;
            string s62 = string.Empty;

            int index = dt.Columns.Count - 5;

            long[] str = new long[index];
            long[] str1 = new long[index];
            d4 = dt.Clone();
            d4.Clear();
            int ijk = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                s1 = dt.Rows[i][0].ToString();
                s2 = dt.Rows[i][1].ToString();
                s3 = dt.Rows[i][2].ToString();
                s31 = dt.Rows[i][3].ToString();
                s32 = dt.Rows[i][4].ToString();

                for (int j = i + 1; j < dt.Rows.Count; j++)
                {
                    s4 = dt.Rows[j][0].ToString();
                    s5 = dt.Rows[j][1].ToString();
                    s6 = dt.Rows[j][2].ToString();
                    s61 = dt.Rows[j][3].ToString();
                    s62 = dt.Rows[j][4].ToString();

                    if (s1 == s4 && s2 == s5 && s3 == s6 && s31 == s61 && s32 == s62)
                    {
                        for (int k = 0; k < index; k++)
                        {
                            try
                            {
                                str1[k] = Convert.ToInt64(dt.Rows[i][k + 5]);
                                str[k] = Convert.ToInt64(dt.Rows[j][k + 5]) + str[k];
                            }
                            catch
                            {
                                continue;
                            }
                        }
                        for (int p = 0; p < dt.Columns.Count; p++)
                        {
                            dt.Rows[j][p] = "";
                        }
                    }
                    else if (s1 != s4 && s2 != s5 && s3 != s6 && s31 != s61 && s32 != s62)
                    {
                        for (int k = 0; k < index; k++)
                        {
                            try
                            {
                                str1[k] = Convert.ToInt64(dt.Rows[i][k + 5]);
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
                d4.Rows.Add();
                for (int l = 0; l < index; l++)
                {
                    d4.Rows[ijk][0] = s1;
                    d4.Rows[ijk][1] = s2;
                    d4.Rows[ijk][2] = s3;
                    d4.Rows[ijk][3] = s31;
                    d4.Rows[ijk][4] = s32;
                    d4.Rows[ijk][l + 5] = str1[l] + str[l];
                }
                for (int k = 0; k < index; k++)
                {
                    str1[k] = 0;
                    str[k] = 0;
                }
                ijk = ijk + 1;
            }
            //DataView dv = dt.DefaultView;
            //dv.Sort = "区分1,物料大类,物料子类,物料编号,物料名称 ASC";  //排序
            //d4= dv.ToTable();
            DataRow[] foundRow;
            foundRow = d4.Select("区分1=''");
            foreach (DataRow row in foundRow)
            {
                d4.Rows.Remove(row);
            }
            return d4;
        }
        //将导入excel文件的datatable精简
        private DataTable StreamlineDT1(DataTable dt)
        {
            DataTable dtnew = new DataTable();
            dt.Columns.Remove("版本");
            dt.Columns.Remove("区分2");
            DataTable dttemp = new DataTable();
            dttemp = dt.Clone();
            dttemp.Clear();
            for (int i = 0; i < (dt.Rows.Count) / 5; i++)
            {
                dttemp.Rows.Add(dt.Rows[i * 5 + 2].ItemArray);

            }
            dtnew = dttemp.Clone();
            dtnew.Clear();
            for (int i = 0; i < dttemp.Rows.Count; i++)
            {
                dtnew.Rows.Add(dttemp.Rows[i].ItemArray);
            }
            return dtnew;
        }
        private DataTable StreamlineDT2(DataTable dt)
        {
            DataTable dtnew = new DataTable();
            DataTable dttemp = new DataTable();
            dttemp = dt.Clone();
            dttemp.Clear();
            for (int i = 0; i < (dt.Rows.Count) / 5; i++)
            {
                dttemp.Rows.Add(dt.Rows[i * 5 + 3].ItemArray);
                dttemp.Rows.Add(dt.Rows[i * 5 + 4].ItemArray);
            }
            dtnew = dttemp.Clone();
            dtnew.Clear();
            for (int i = 0; i < dttemp.Rows.Count / 2; i++)
            {
                dtnew.Rows.Add(dttemp.Rows[2 * i + 0].ItemArray);
                dtnew.Rows.Add(dttemp.Rows[2 * i + 1].ItemArray);
            }
            return dtnew;
        }
        //导入EXCEL文件
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
                        DataColumn datacolum;
                        try
                        {
                            datacolum = new DataColumn(headrow.GetCell(i).StringCellValue);
                        }
                        catch
                        {
                            datacolum = new DataColumn(headrow.GetCell(i).DateCellValue.ToString("yyyy-MM-dd"));
                        }
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
        //读文本记录时间
        private string ReadDateTime()
        {
            if (!File.Exists("data/time.txt"))
            {
                return "1900-01-01 00:00:00";
            }
            else
            {
                FileStream fs = new FileStream("data/time.txt", FileMode.Open, FileAccess.Read);
                StreamReader sd = new StreamReader(fs);
                string str = sd.ReadLine();
                fs.Close();
                return str;
            }
        }
        private void WriteDateTime()
        {
            FileStream fs = new FileStream("data/time.txt", FileMode.Create, FileAccess.Write);
            StreamWriter sr = new StreamWriter(fs);
            sr.WriteLine(DateTime.Now.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"));
            sr.Close();
            fs.Close();
        }
        //插入工序基本信息表
        private void InsertGXData(DataTable dt, bool b)
        {
            SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
            string s = connect.GetConnectStr("ERP");
            string connectionString = s;

            if (b == true)
            {
                //根据比较结果更新本地数据库
                string connStr = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                using (SQLiteConnection con = new SQLiteConnection(connStr))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "INSERT INTO BasicProcessInformation (CPBH,ZBXH_BJBH,BJMC,SJC,SJK,ZMYS,FMYS,FSMC,DL,GXBH,GBLB,GXMC,BHCS) VALUES(@CPBH,@ZBXH_BJBH,@BJMC,@SJC,@SJK,@ZMYS,@FMYS,@FSMC,@DL,@GXBH,@GBLB,@GXMC,@BHCS)";
                    for (int n = 0; n < dt.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@ZBXH_BJBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@ZMYS", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@FMYS", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@FSMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GBLB", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BHCS", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@ZBXH_BJBH"].Value = dt.Rows[n]["部件编号"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt.Rows[n]["部件名称"].ToString();
                        cmd.Parameters["@SJC"].Value = dt.Rows[n]["上机长"].ToString();
                        cmd.Parameters["@SJK"].Value = dt.Rows[n]["上机宽"].ToString();
                        cmd.Parameters["@ZMYS"].Value = dt.Rows[n]["正面颜色"].ToString();
                        cmd.Parameters["@FMYS"].Value = dt.Rows[n]["反面颜色"].ToString();
                        cmd.Parameters["@FSMC"].Value = dt.Rows[n]["印刷方式"].ToString();
                        cmd.Parameters["@DL"].Value = dt.Rows[n]["得率"].ToString();
                        cmd.Parameters["@GXBH"].Value = dt.Rows[n]["工序编码"].ToString();
                        cmd.Parameters["@GBLB"].Value = dt.Rows[n]["工序类别"].ToString();
                        cmd.Parameters["@GXMC"].Value = dt.Rows[n]["工序名称"].ToString();
                        cmd.Parameters["@BHCS"].Value = dt.Rows[n]["变化系数"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
            }
            else
            {
                //查询sql server数据库更新本地数据库       
                DataTable dt1 = new DataTable();
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "ProcessOfConversion2";
                        SqlParameter[] para ={
                                       new SqlParameter("@times",SqlDbType.DateTime)};
                        para[0].Value = ReadDateTime();
                        try
                        {
                            cmd.CommandTimeout = 60 * 60 * 1000;
                            cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(dt1);
                            SqlHelper.GetConnection().Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            SqlHelper.GetConnection().Close();
                        }
                    }
                }
                ox = dt1.Rows.Count;
                //删除数据
                string connStr = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                using (SQLiteConnection con = new SQLiteConnection(connStr))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "delete from  BasicProcessInformation where CPBH=@CPBH";
                    for (int n = 0; n < dt1.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt1.Rows[n]["产品编号"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
                //新增数据
                using (SQLiteConnection con = new SQLiteConnection(connStr))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "INSERT INTO BasicProcessInformation (CPBH,CPMC,ZBXH_BJBH,BJMC,SJC,SJK,ZMYS,FMYS,FSMC,DL,GXBH,GBLB,GXMC,BHCS) VALUES(@CPBH,@CPMC,@ZBXH_BJBH,@BJMC,@SJC,@SJK,@ZMYS,@FMYS,@FSMC,@DL,@GXBH,@GBLB,@GXMC,@BHCS)";
                    for (int n = 0; n < dt1.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@ZBXH_BJBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@ZMYS", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@FMYS", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@FSMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GBLB", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@GXMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BHCS", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt1.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@CPMC"].Value = dt1.Rows[n]["产品名称"].ToString();
                        cmd.Parameters["@ZBXH_BJBH"].Value = dt1.Rows[n]["部件编号"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt1.Rows[n]["部件名称"].ToString();
                        cmd.Parameters["@SJC"].Value = dt1.Rows[n]["上机长"].ToString();
                        cmd.Parameters["@SJK"].Value = dt1.Rows[n]["上机宽"].ToString();
                        cmd.Parameters["@ZMYS"].Value = dt1.Rows[n]["正面颜色"].ToString();
                        cmd.Parameters["@FMYS"].Value = dt1.Rows[n]["反面颜色"].ToString();
                        cmd.Parameters["@FSMC"].Value = dt1.Rows[n]["印刷方式"].ToString();
                        cmd.Parameters["@DL"].Value = dt1.Rows[n]["得率"].ToString();
                        cmd.Parameters["@GXBH"].Value = dt1.Rows[n]["工序编码"].ToString();
                        cmd.Parameters["@GBLB"].Value = dt1.Rows[n]["工序类别"].ToString();
                        cmd.Parameters["@GXMC"].Value = dt1.Rows[n]["工序名称"].ToString();
                        cmd.Parameters["@BHCS"].Value = dt1.Rows[n]["变化系数"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
            }
        }
        //插入物料基本信息表
        private void InsertWLData(DataTable dt, bool b)
        {
            SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
            string s = connect.GetConnectStr("ERP");
            string connectionString = s;

            if (b == true)
            {
                //根据比较结果更新本地数据库
                string connStr1 = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                using (SQLiteConnection con = new SQLiteConnection(connStr1))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "INSERT INTO BasicMaterials (CPBH,CPMC,BJMC,SJC,SJK,DL,WLDL,WLZL,WLBH,WLMC,BZYL,WLDW) VALUES(@CPBH,@CPMC,@BJMC,@SJC,@SJK,@DL,@WLDL,@WLZL,@WLBH,@WLMC,@BZYL,@WLDW)";
                    for (int n = 0; n < dt.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLDL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLZL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BZYL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLDW", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@CPMC"].Value = dt.Rows[n]["产品名称"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt.Rows[n]["部件编号"].ToString();
                        cmd.Parameters["@SJC"].Value = dt.Rows[n]["上机长"].ToString();
                        cmd.Parameters["@SJK"].Value = dt.Rows[n]["上机宽"].ToString();
                        cmd.Parameters["@DL"].Value = dt.Rows[n]["得率"].ToString();
                        cmd.Parameters["@WLDL"].Value = dt.Rows[n]["物料大类"].ToString();
                        cmd.Parameters["@WLZL"].Value = dt.Rows[n]["物料小类"].ToString();
                        cmd.Parameters["@WLBH"].Value = dt.Rows[n]["物料编号"].ToString();
                        cmd.Parameters["@WLMC"].Value = dt.Rows[n]["物料名称"].ToString();
                        cmd.Parameters["@BZYL"].Value = dt.Rows[n]["标准用量"].ToString();
                        cmd.Parameters["@WLDL"].Value = dt.Rows[n]["单位名称"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
            }
            else
            {
                //查询sql server数据库更新本地数据库       
                DataTable dt2 = new DataTable();
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "MaterialQuery1";
                        SqlParameter[] para = { new SqlParameter("@times", SqlDbType.DateTime) };
                        para[0].Value = ReadDateTime();
                        try
                        {
                            cmd.CommandTimeout = 60 * 60 * 1000;
                            cmd.Parameters.AddRange(para);// 将参数加入命令对象  
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(dt2);
                            SqlHelper.GetConnection().Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            SqlHelper.GetConnection().Close();
                        }
                    }
                }
                oy = dt2.Rows.Count;
                string connStr1 = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
                //删除数据
                using (SQLiteConnection con = new SQLiteConnection(connStr1))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "delete from  BasicMaterials where CPBH=@CPBH ";
                    for (int n = 0; n < dt2.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt2.Rows[n]["产品编号"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
                //插入数据
                using (SQLiteConnection con = new SQLiteConnection(connStr1))
                {
                    con.Open();
                    DbTransaction trans = con.BeginTransaction();//开始事务       
                    SQLiteCommand cmd = new SQLiteCommand(con);
                    cmd.CommandText = "INSERT INTO BasicMaterials (CPBH,CPMC,BJMC,SJC,SJK,DL,WLDL,WLZL,WLBH,WLMC,BZYL,WLDW) VALUES(@CPBH,@CPMC,@BJMC,@SJC,@SJK,@DL,@WLDL,@WLZL,@WLBH,@WLMC,@BZYL,@WLDW)";
                    for (int n = 0; n < dt2.Rows.Count; n++)
                    {
                        cmd.Parameters.Add(new SQLiteParameter("@CPBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BJMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@SJK", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@DL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLDL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLZL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLBH", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLMC", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@BZYL", DbType.String));
                        cmd.Parameters.Add(new SQLiteParameter("@WLDW", DbType.String));
                        cmd.Parameters["@CPBH"].Value = dt2.Rows[n]["产品编号"].ToString();
                        cmd.Parameters["@CPMC"].Value = dt2.Rows[n]["产品名称"].ToString();
                        cmd.Parameters["@BJMC"].Value = dt2.Rows[n]["部件编号"].ToString();
                        cmd.Parameters["@SJC"].Value = dt2.Rows[n]["上机长"].ToString();
                        cmd.Parameters["@SJK"].Value = dt2.Rows[n]["上机宽"].ToString();
                        cmd.Parameters["@DL"].Value = dt2.Rows[n]["得率"].ToString();
                        cmd.Parameters["@WLDL"].Value = dt2.Rows[n]["物料大类"].ToString();
                        cmd.Parameters["@WLZL"].Value = dt2.Rows[n]["物料小类"].ToString();
                        cmd.Parameters["@WLBH"].Value = dt2.Rows[n]["物料编号"].ToString();
                        cmd.Parameters["@WLMC"].Value = dt2.Rows[n]["物料名称"].ToString();
                        cmd.Parameters["@BZYL"].Value = dt2.Rows[n]["标准用量"].ToString();
                        cmd.Parameters["@WLDW"].Value = dt2.Rows[n]["单位名称"].ToString();
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();//提交事务  
                    con.Close();
                }
            }
        }

        #endregion

        #region 连接两个表
        public static DataTable Join(DataTable First, DataTable Second, DataColumn[] FJC, DataColumn[] SJC)
        {
            DataTable table1 = new DataTable("Join");
            using (DataSet ds1 = new DataSet())
            {
                ds1.Tables.AddRange(new DataTable[] { First.Copy(), Second.Copy() });
                DataColumn[] First_columns = new DataColumn[FJC.Length];
                for (int i = 0; i < First_columns.Length; i++)
                {
                    First_columns[i] = ds1.Tables[0].Columns[FJC[i].ColumnName];
                }
                DataColumn[] Second_columns = new DataColumn[SJC.Length];
                for (int i = 0; i < Second_columns.Length; i++)
                {
                    Second_columns[i] = ds1.Tables[1].Columns[SJC[i].ColumnName];
                }
                DataRelation r = new DataRelation(string.Empty, First_columns, Second_columns, false);
                ds1.Relations.Add(r);

                for (int i = 0; i < First.Columns.Count; i++)
                {
                    table1.Columns.Add(First.Columns[i].ColumnName, First.Columns[i].DataType);
                }

                for (int i = 0; i < Second.Columns.Count; i++)
                {

                    //看看有没有重复的列，如果有在第二个DataTable的Column的列明后加_Second
                    if (!table1.Columns.Contains(Second.Columns[i].ColumnName))
                        table1.Columns.Add(Second.Columns[i].ColumnName, Second.Columns[i].DataType);
                    else
                        table1.Columns.Add(Second.Columns[i].ColumnName + "_1", Second.Columns[i].DataType);
                }
                table1.BeginLoadData();

                int itable2Colomns = 0;
                if (ds1.Tables[1].Rows.Count > 0)
                {
                    itable2Colomns = ds1.Tables[1].Rows[0].ItemArray.Length;
                }

                foreach (DataRow firstrow in ds1.Tables[0].Rows)
                {
                    //得到行的数据
                    DataRow[] childrows = firstrow.GetChildRows(r);//第二个表关联的行
                    if (childrows != null && childrows.Length > 0)
                    {
                        object[] parentarray = firstrow.ItemArray;
                        foreach (DataRow secondrow in childrows)
                        {
                            object[] secondarray = secondrow.ItemArray;
                            object[] joinarray = new object[parentarray.Length + secondarray.Length];
                            System.Array.Copy(parentarray, 0, joinarray, 0, parentarray.Length);
                            System.Array.Copy(secondarray, 0, joinarray, parentarray.Length, secondarray.Length);
                            table1.LoadDataRow(joinarray, true);
                        }
                    }
                    else//如果有外连接(Left Join)添加这部分代码
                    {
                        object[] table1array = firstrow.ItemArray;//Table1
                        object[] table2array = new object[itable2Colomns];
                        object[] joinarray = new object[table1array.Length + itable2Colomns];
                        System.Array.Copy(table1array, 0, joinarray, 0, table1array.Length);
                        System.Array.Copy(table2array, 0, joinarray, table1array.Length, itable2Colomns);
                        table1.LoadDataRow(joinarray, true);
                        DataColumn[] dc = new DataColumn[2];
                        dc[0] = new DataColumn("");
                    }
                }
                table1.EndLoadData();
            }
            return table1;
        }
        //将两个结构相同的datatable合并
        private DataTable adddatatable(DataTable DataTable1, DataTable DataTable2)
        {
            DataTable newDataTable = DataTable1.Clone();

            object[] obj = new object[newDataTable.Columns.Count];
            for (int i = 0; i < DataTable1.Rows.Count; i++)
            {
                DataTable1.Rows[i].ItemArray.CopyTo(obj, 0);
                newDataTable.Rows.Add(obj);
            }

            for (int i = 0; i < DataTable2.Rows.Count; i++)
            {
                DataTable2.Rows[i].ItemArray.CopyTo(obj, 0);
                newDataTable.Rows.Add(obj);
            }
            return newDataTable;
        }
        //根据行数切割表格
        public DataSet SplitDataTable(DataTable originalTab, int rowsNum)
        {
            //获取所需创建的表数量
            int tableNum = originalTab.Rows.Count / rowsNum;

            //获取数据余数
            int remainder = originalTab.Rows.Count % rowsNum;

            DataSet ds = new DataSet();

            //如果只需要创建1个表，直接将原始表存入DataSet
            if (tableNum == 0)
            {
                ds.Tables.Add(originalTab);
            }
            else
            {
                DataTable[] tableSlice = new DataTable[tableNum];

                //Save orginal columns into new table.            
                for (int c = 0; c < tableNum; c++)
                {
                    tableSlice[c] = new DataTable();
                    foreach (DataColumn dc in originalTab.Columns)
                    {
                        tableSlice[c].Columns.Add(dc.ColumnName, dc.DataType);
                    }
                }
                //Import Rows
                for (int i = 0; i < tableNum; i++)
                {
                    // if the current table is not the last one
                    if (i != tableNum - 1)
                    {
                        for (int j = i * rowsNum; j < ((i + 1) * rowsNum); j++)
                        {
                            tableSlice[i].ImportRow(originalTab.Rows[j]);
                        }
                    }
                    else
                    {
                        for (int k = i * rowsNum; k < ((i + 1) * rowsNum + remainder); k++)
                        {
                            tableSlice[i].ImportRow(originalTab.Rows[k]);
                        }
                    }
                }

                //add all tables into a dataset                
                foreach (DataTable dt in tableSlice)
                {
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }
        /// <summary>
        /// 连接两个表
        /// </summary>
        /// <param name="First"></param>
        /// <param name="Second"></param>
        /// <param name="FJC"></param>
        /// <param name="SJC"></param>
        /// <returns></returns>
        public static DataTable JoinTwoTable(DataTable First, DataTable Second, string FJC, string SJC)
        {
            return JoinTwoTable(First, Second, new DataColumn[] { First.Columns[FJC] }, new DataColumn[] { First.Columns[SJC] });
        }
        /// <summary>
        /// 连接两个表
        /// </summary>
        /// <param name="First"></param>
        /// <param name="Second"></param>
        /// <param name="FJC"></param>
        /// <param name="SJC"></param>
        /// <returns></returns>
        protected static DataTable JoinTwoTable(DataTable First, DataTable Second, DataColumn FJC, DataColumn SJC)
        {
            return JoinTwoTable(First, Second, new DataColumn[] { FJC }, new DataColumn[] { SJC });
        }
        /// <summary>
        /// 连接两个Table
        /// </summary>
        /// <param name="First"></param>
        /// <param name="Second"></param>
        /// <param name="FJC"></param>
        /// <param name="SJC"></param>
        /// <returns></returns>
        protected static DataTable JoinTwoTable(DataTable First, DataTable Second, DataColumn[] FJC, DataColumn[] SJC)
        {
            //创建一个新的DataTable
            DataTable table = new DataTable("Join");
            using (DataSet ds = new DataSet())
            {
                //把DataTable Copy到DataSet中
                ds.Tables.AddRange(new DataTable[] { First.Copy(), Second.Copy() });
                DataColumn[] parentcolumns = new DataColumn[FJC.Length];
                for (int i = 0; i < parentcolumns.Length; i++)
                {
                    parentcolumns[i] = ds.Tables[0].Columns[FJC[i].ColumnName];
                }
                DataColumn[] childcolumns = new DataColumn[SJC.Length];
                for (int i = 0; i < childcolumns.Length; i++)
                {
                    childcolumns[i] = ds.Tables[1].Columns[SJC[i].ColumnName];
                }
                //创建关联
                DataRelation r = new DataRelation(string.Empty, parentcolumns, childcolumns, false);
                ds.Relations.Add(r);
                //为关联表创建列
                for (int i = 0; i < First.Columns.Count; i++)
                {
                    table.Columns.Add(First.Columns[i].ColumnName, First.Columns[i].DataType);
                }
                for (int i = 0; i < Second.Columns.Count; i++)
                {
                    //看看有没有重复的列，如果有在第二个DataTable的Column的列明后加_Second
                    if (!table.Columns.Contains(Second.Columns[i].ColumnName))
                        table.Columns.Add(Second.Columns[i].ColumnName, Second.Columns[i].DataType);
                    else
                        table.Columns.Add(Second.Columns[i].ColumnName + "_Second", Second.Columns[i].DataType);
                }
                table.BeginLoadData();
                foreach (DataRow firstrow in ds.Tables[0].Rows)
                {
                    //得到行的数据
                    DataRow[] childrows = firstrow.GetChildRows(r);
                    if (childrows != null && childrows.Length > 0)
                    {
                        object[] parentarray = firstrow.ItemArray;
                        foreach (DataRow secondrow in childrows)
                        {
                            object[] secondarray = secondrow.ItemArray;
                            object[] joinarray = new object[parentarray.Length + secondarray.Length];
                            Array.Copy(parentarray, 0, joinarray, 0, parentarray.Length);
                            Array.Copy(secondarray, 0, joinarray, parentarray.Length, secondarray.Length);
                            table.LoadDataRow(joinarray, true);
                        }
                    }
                }
                table.EndLoadData();
            }
            return table;
        }

        #endregion

        #region 新方法
        private string ConvertStr(object o)
        {
            try
            {
                return Convert.ToString(o);
            }
            catch
            {
                return "";
            }
        }
        private DataTable n1()
        {
            DataTable dtrr = new DataTable();
            string sql = "SELECT d1.XH AS 序号,	d1.KHBM AS 客户名称,	d1.CPBM AS 产品编码,	D1.CPMC AS 产品名称,	d1.DDYLH AS 订单余量合计,	d1.SHSLH AS 送货数量合计,	d1.kc AS 库存,	D1.YYS AS 已印刷,	D1.dys AS 待印刷,	d1.WQWZH AS 往期未转化为施工单的订单,	d1.PDSFCC AS 判断是否超产需要扣数,	D1.QUY AS 区分1,	d1.JHTS AS 交货天数,	m.GXBH AS 工序编码,	m.GBLB AS 工序类别,	M.GXMC AS 工序名称,	( d1.D1 * m.DL * m.BHCS ) AS D1,	( d1.D2 * m.DL * m.BHCS ) AS D2,	( d1.D3 * m.DL * m.BHCS ) AS D3,	( d1.D4 * m.DL * m.BHCS ) AS D4,	( d1.D5 * m.DL * m.BHCS ) AS D5,	( d1.D6 * m.DL * m.BHCS ) AS D6,	( d1.D7 * m.DL * m.BHCS ) AS D7,	( d1.D8 * m.DL * m.BHCS ) AS D8,	( d1.D9 * m.DL * m.BHCS ) AS D9,	( d1.D10 * m.DL * m.BHCS ) AS D10,	( d1.D11 * m.DL * m.BHCS ) AS D11,	( d1.D12 * m.DL * m.BHCS ) AS D12,	( d1.D13 * m.DL * m.BHCS ) AS D13,	( d1.D14 * m.DL * m.BHCS ) AS D14,	( d1.D15 * m.DL * m.BHCS ) AS D15,	( d1.D16 * m.DL * m.BHCS ) AS D16,	( d1.D17 * m.DL * m.BHCS ) AS D17,	( d1.D18 * m.DL * m.BHCS ) AS D18,	( d1.D19 * m.DL * m.BHCS ) AS D19,	( d1.D20 * m.DL * m.BHCS ) AS D20,	( d1.D21 * m.DL * m.BHCS ) AS D21,	( d1.D22 * m.DL * m.BHCS ) AS D22,	( d1.D23 * m.DL * m.BHCS ) AS D23,	( d1.D24 * m.DL * m.BHCS ) AS D24,	( d1.D25 * m.DL * m.BHCS ) AS D25,	( d1.D26 * m.DL * m.BHCS ) AS D26,	( d1.D27 * m.DL * m.BHCS ) AS D27,	( d1.D28 * m.DL * m.BHCS ) AS D28,	( d1.D29 * m.DL * m.BHCS ) AS D29,	( d1.D30 * m.DL * m.BHCS ) AS D30,	( d1.D31 * m.DL * m.BHCS ) AS D31,	( d1.D32 * m.DL * m.BHCS ) AS D32,	( d1.D33 * m.DL * m.BHCS ) AS D33,	( d1.D34 * m.DL * m.BHCS ) AS D34,	( d1.D35 * m.DL * m.BHCS ) AS D35 FROM BPTEMP D1 left JOIN BasicProcessInformation M ON m.CPBH = d1.CPBM";
            dtrr = Tools.Common.SQLiteHelper.ExecuteDatatable(sql);
            for (int i = 0; i < dtrr.Rows.Count; i++)
            {
                if (dtrr.Rows[i]["工序编码"].ToString() == "" && dtrr.Rows[i]["工序类别"].ToString() == "" && dtrr.Rows[i]["工序名称"].ToString() == "")
                {
                    dtrr.Rows[i]["产品名称"] = dtrr.Rows[i]["产品名称"] + "该产品档案属于未审核状态";
                }
            }
            return dtrr;
        }
        private DataTable n5()
        {
            DataTable dtrr = new DataTable();
            string sql = "SELECT	d1.XH AS 序号,	d1.KHBM AS 客户名称,	d1.CPBM AS 产品编码,	D1.CPMC AS 产品名称,	d1.DDYLH AS 订单余量合计,	d1.SHSLH AS 送货数量合计,	d1.kc AS 库存,	D1.YYS AS 已印刷,	D1.dys AS 待印刷,	d1.WQWZH AS 往期未转化为施工单的订单,	d1.PDSFCC AS 判断是否超产需要扣数,	D1.QUY AS 区分1,	d1.JHTS AS 交货天数,	m.GXBH AS 工序编码,	m.GBLB AS 工序类别,	M.GXMC AS 工序名称,	( d1.D1 * m.DL * m.BHCS ) AS D1,	( d1.D2 * m.DL * m.BHCS ) AS D2,	( d1.D3 * m.DL * m.BHCS ) AS D3,	( d1.D4 * m.DL * m.BHCS ) AS D4,	( d1.D5 * m.DL * m.BHCS ) AS D5,	( d1.D6 * m.DL * m.BHCS ) AS D6,	( d1.D7 * m.DL * m.BHCS ) AS D7,	( d1.D8 * m.DL * m.BHCS ) AS D8,	( d1.D9 * m.DL * m.BHCS ) AS D9,	( d1.D10 * m.DL * m.BHCS ) AS D10,	( d1.D11 * m.DL * m.BHCS ) AS D11,	( d1.D12 * m.DL * m.BHCS ) AS D12,	( d1.D13 * m.DL * m.BHCS ) AS D13,	( d1.D14 * m.DL * m.BHCS ) AS D14,	( d1.D15 * m.DL * m.BHCS ) AS D15,	( d1.D16 * m.DL * m.BHCS ) AS D16,	( d1.D17 * m.DL * m.BHCS ) AS D17,	( d1.D18 * m.DL * m.BHCS ) AS D18,	( d1.D19 * m.DL * m.BHCS ) AS D19,	( d1.D20 * m.DL * m.BHCS ) AS D20,	( d1.D21 * m.DL * m.BHCS ) AS D21,	( d1.D22 * m.DL * m.BHCS ) AS D22,	( d1.D23 * m.DL * m.BHCS ) AS D23,	( d1.D24 * m.DL * m.BHCS ) AS D24,	( d1.D25 * m.DL * m.BHCS ) AS D25,	( d1.D26 * m.DL * m.BHCS ) AS D26,	( d1.D27 * m.DL * m.BHCS ) AS D27,	( d1.D28 * m.DL * m.BHCS ) AS D28,	( d1.D29 * m.DL * m.BHCS ) AS D29,	( d1.D30 * m.DL * m.BHCS ) AS D30,	( d1.D31 * m.DL * m.BHCS ) AS D31,	( d1.D32 * m.DL * m.BHCS ) AS D32,	( d1.D33 * m.DL * m.BHCS ) AS D33,	( d1.D34 * m.DL * m.BHCS ) AS D34,	( d1.D35 * m.DL * m.BHCS ) AS D35 FROM BPTEMP D1 left JOIN BasicProcessInformation M ON m.CPBH = d1.CPBM";
            dtrr = Tools.Common.SQLiteHelper.ExecuteDatatable(sql);
            return dtrr;
        }
        private DataTable n2()
        {
            DataTable dtrr = new DataTable();
            string sql = "select 	 m.GXBH as 工序编码,	 m.GBLB as 工序类别,	 M.GXMC AS 工序名称,	 sum(d1.D1*m.DL*m.BHCS) as D1,	 sum(d1.D2*m.DL*m.BHCS) as D2 ,	 sum(d1.D3*m.DL*m.BHCS) as D3 ,	 sum(d1.D4*m.DL*m.BHCS) as D4 ,	 sum(d1.D5*m.DL*m.BHCS) as D5 ,	 sum(d1.D6*m.DL*m.BHCS) as D6 ,	 sum(d1.D7*m.DL*m.BHCS) as D7 ,	 sum(d1.D8*m.DL*m.BHCS) as D8 ,	 sum(d1.D9*m.DL*m.BHCS) as D9 ,	 sum(d1.D10*m.DL*m.BHCS) as D10,	 sum(d1.D11*m.DL*m.BHCS) as D11,	 sum(d1.D12*m.DL*m.BHCS) as D12,	 sum(d1.D13*m.DL*m.BHCS) as D13,	 sum(d1.D14*m.DL*m.BHCS) as D14,	 sum(d1.D15*m.DL*m.BHCS) as D15,	 sum(d1.D16*m.DL*m.BHCS) as D16,	 sum(d1.D17*m.DL*m.BHCS) as D17,	 sum(d1.D18*m.DL*m.BHCS) as D18,	 sum(d1.D19*m.DL*m.BHCS) as D19,	 sum(d1.D20*m.DL*m.BHCS) as D20,	 sum(d1.D21*m.DL*m.BHCS) as D21,	 sum(d1.D22*m.DL*m.BHCS) as D22,	 sum(d1.D23*m.DL*m.BHCS) as D23,	 sum(d1.D24*m.DL*m.BHCS) as D24,	 sum(d1.D25*m.DL*m.BHCS) as D25,	 sum(d1.D26*m.DL*m.BHCS) as D26,	 sum(d1.D27*m.DL*m.BHCS) as D27,	 sum(d1.D28*m.DL*m.BHCS) as D28,	 sum(d1.D29*m.DL*m.BHCS) as D29,	 sum(d1.D30*m.DL*m.BHCS) as D30,	 sum(d1.D31*m.DL*m.BHCS) as D31,	 sum(d1.D32*m.DL*m.BHCS) as D32,	 sum(d1.D33*m.DL*m.BHCS) as D33,	 sum(d1.D34*m.DL*m.BHCS) as D34,	 sum(d1.D35*m.DL*m.BHCS) as D35 from BasicProcessInformation M inner join BPTEMP D1  on m.CPBH=d1.CPBM  group by m.GXBH,m.GBLB, M.GXMC";
            dtrr = Tools.Common.SQLiteHelper.ExecuteDatatable(sql);
            for (int i = 0; i < dtrr.Rows.Count; i++)
            {
                if (dtrr.Rows[i]["工序编码"].ToString() == "")
                {
                    dtrr.Rows.RemoveAt(i);
                }
            }
            return dtrr;
        }
        private DataTable n3()
        {
            DataTable dtrr = new DataTable();
            string sql = "SELECT	d1.XH AS 序号,	d1.KHBM AS 客户名称,	d1.CPBM AS 产品编码,	D1.CPMC AS 产品名称,	d1.DDYLH AS 订单余量合计,	d1.SHSLH AS 送货数量合计,	d1.kc AS 库存,	D1.YYS AS 已印刷,	D1.dys AS 待印刷,	d1.WQWZH AS 往期未转化为施工单的订单,	d1.PDSFCC AS 判断是否超产需要扣数,	D1.QUY AS 区分1,	d1.JHTS AS 交货天数,	M.WLDL AS 物料大类,	M.WLZL AS 物料子类,	M.WLBH AS 物料编码,	M.WLMC AS 物料名称,	m.WLDW as 物料单位,	sum( d1.D1 * m.DL ) AS D1,	sum( d1.D2 * m.DL ) AS D2,	sum( d1.D3 * m.DL ) AS D3,	sum( d1.D4 * m.DL ) AS D4,	sum( d1.D5 * m.DL ) AS D5,	sum( d1.D6 * m.DL ) AS D6,	sum( d1.D7 * m.DL ) AS D7,	sum( d1.D8 * m.DL ) AS D8,	sum( d1.D9 * m.DL ) AS D9,	sum( d1.D10 * m.DL ) AS D10,	sum( d1.D11 * m.DL ) AS D11,	sum( d1.D12 * m.DL ) AS D12,	sum( d1.D13 * m.DL ) AS D13,	sum( d1.D14 * m.DL ) AS D14,	sum( d1.D15 * m.DL ) AS D15,	sum( d1.D16 * m.DL ) AS D16,	sum( d1.D17 * m.DL ) AS D17,	sum( d1.D18 * m.DL ) AS D18,	sum( d1.D19 * m.DL ) AS D19,	sum( d1.D20 * m.DL ) AS D20,	sum( d1.D21 * m.DL ) AS D21,	sum( d1.D22 * m.DL ) AS D22,	sum( d1.D23 * m.DL ) AS D23,	sum( d1.D24 * m.DL ) AS D24,	sum( d1.D25 * m.DL ) AS D25,	sum( d1.D26 * m.DL ) AS D26,	sum( d1.D27 * m.DL ) AS D27,	sum( d1.D28 * m.DL ) AS D28,	sum( d1.D29 * m.DL ) AS D29,	sum( d1.D30 * m.DL ) AS D30,	sum( d1.D31 * m.DL ) AS D31,	sum( d1.D32 * m.DL ) AS D32,	sum( d1.D33 * m.DL ) AS D33,	sum( d1.D34 * m.DL ) AS D34,	sum( d1.D35 * m.DL ) AS D35 FROM	BMTEMP D1	LEFT JOIN BasicMaterials M ON m.CPBH = d1.CPBM where d1.quy='订单需求' GROUP BY	d1.XH,	d1.KHBM,	d1.CPBM,	D1.CPMC,	d1.DDYLH,	d1.SHSLH,	d1.kc,	D1.YYS,	D1.dys,	d1.WQWZH,	d1.PDSFCC,	D1.QUY,	d1.JHTS,	M.WLDL,	M.WLZL,	M.WLBH,	M.WLMC,	m.WLDW";
            DataTable dtrr1 = Tools.Common.SQLiteHelper.ExecuteDatatable(sql);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
            string sql1 = "SELECT	d1.XH AS 序号,	d1.KHBM AS 客户名称,	d1.CPBM AS 产品编码,	D1.CPMC AS 产品名称,	d1.DDYLH AS 订单余量合计,	d1.SHSLH AS 送货数量合计,	d1.kc AS 库存,	D1.YYS AS 已印刷,	D1.dys AS 待印刷,	d1.WQWZH AS 往期未转化为施工单的订单,	d1.PDSFCC AS 判断是否超产需要扣数,	D1.QUY AS 区分1,	d1.JHTS AS 交货天数,	M.WLDL AS 物料大类,	M.WLZL AS 物料子类,	M.WLBH AS 物料编码,	M.WLMC AS 物料名称,	m.WLDW as 物料单位,	sum( d1.D1 * m.DL ) AS D1,	sum( d1.D2 * m.DL ) AS D2,	sum( d1.D3 * m.DL ) AS D3,	sum( d1.D4 * m.DL ) AS D4,	sum( d1.D5 * m.DL ) AS D5,	sum( d1.D6 * m.DL ) AS D6,	sum( d1.D7 * m.DL ) AS D7,	sum( d1.D8 * m.DL ) AS D8,	sum( d1.D9 * m.DL ) AS D9,	sum( d1.D10 * m.DL ) AS D10,	sum( d1.D11 * m.DL ) AS D11,	sum( d1.D12 * m.DL ) AS D12,	sum( d1.D13 * m.DL ) AS D13,	sum( d1.D14 * m.DL ) AS D14,	sum( d1.D15 * m.DL ) AS D15,	sum( d1.D16 * m.DL ) AS D16,	sum( d1.D17 * m.DL ) AS D17,	sum( d1.D18 * m.DL ) AS D18,	sum( d1.D19 * m.DL ) AS D19,	sum( d1.D20 * m.DL ) AS D20,	sum( d1.D21 * m.DL ) AS D21,	sum( d1.D22 * m.DL ) AS D22,	sum( d1.D23 * m.DL ) AS D23,	sum( d1.D24 * m.DL ) AS D24,	sum( d1.D25 * m.DL ) AS D25,	sum( d1.D26 * m.DL ) AS D26,	sum( d1.D27 * m.DL ) AS D27,	sum( d1.D28 * m.DL ) AS D28,	sum( d1.D29 * m.DL ) AS D29,	sum( d1.D30 * m.DL ) AS D30,	sum( d1.D31 * m.DL ) AS D31,	sum( d1.D32 * m.DL ) AS D32,	sum( d1.D33 * m.DL ) AS D33,	sum( d1.D34 * m.DL ) AS D34,	sum( d1.D35 * m.DL ) AS D35 FROM	BMTEMP D1	LEFT JOIN BasicMaterials M ON m.CPBH = d1.CPBM where d1.quy='预测需求' 	 GROUP BY	d1.XH,	d1.KHBM,	d1.CPBM,	D1.CPMC,	d1.DDYLH,	d1.SHSLH,	d1.kc,	D1.YYS,	D1.dys,	d1.WQWZH,	d1.PDSFCC,	D1.QUY,	d1.JHTS,	M.WLDL,	M.WLZL,	M.WLBH,	M.WLMC,	m.WLDW";
            DataTable dtrr2 = Tools.Common.SQLiteHelper.ExecuteDatatable(sql1);
            dtrr = adddatatable(dtrr1, dtrr2);
            for (int i = 0; i < dtrr.Rows.Count; i++)
            {
                if (dtrr.Rows[i]["物料大类"].ToString() == "" && dtrr.Rows[i]["物料子类"].ToString() == "" && dtrr.Rows[i]["物料编码"].ToString() == "" && dtrr.Rows[i]["物料名称"].ToString() == "")
                {
                    dtrr.Rows[i]["产品名称"] = dtrr.Rows[i]["产品名称"] + "该产品档案属于未审核状态";
                }
            }
            index1 = dtrr.Rows.Count/2;
            return dtrr;
        }
        private DataTable n6()
        {
            DataTable dtrr = new DataTable();
            string sql = "SELECT	d1.XH AS 序号,	d1.KHBM AS 客户名称,	d1.CPBM AS 产品编码,	D1.CPMC AS 产品名称,	d1.DDYLH AS 订单余量合计,	d1.SHSLH AS 送货数量合计,	d1.kc AS 库存,	D1.YYS AS 已印刷,	D1.dys AS 待印刷,	d1.WQWZH AS 往期未转化为施工单的订单,	d1.PDSFCC AS 判断是否超产需要扣数,	D1.QUY AS 区分1,	d1.JHTS AS 交货天数,	M.WLDL AS 物料大类,	M.WLZL AS 物料子类,	M.WLBH AS 物料编码,	M.WLMC AS 物料名称,	m.WLDW as 物料单位,	sum( d1.D1 * m.DL ) AS D1,	sum( d1.D2 * m.DL ) AS D2,	sum( d1.D3 * m.DL ) AS D3,	sum( d1.D4 * m.DL ) AS D4,	sum( d1.D5 * m.DL ) AS D5,	sum( d1.D6 * m.DL ) AS D6,	sum( d1.D7 * m.DL ) AS D7,	sum( d1.D8 * m.DL ) AS D8,	sum( d1.D9 * m.DL ) AS D9,	sum( d1.D10 * m.DL ) AS D10,	sum( d1.D11 * m.DL ) AS D11,	sum( d1.D12 * m.DL ) AS D12,	sum( d1.D13 * m.DL ) AS D13,	sum( d1.D14 * m.DL ) AS D14,	sum( d1.D15 * m.DL ) AS D15,	sum( d1.D16 * m.DL ) AS D16,	sum( d1.D17 * m.DL ) AS D17,	sum( d1.D18 * m.DL ) AS D18,	sum( d1.D19 * m.DL ) AS D19,	sum( d1.D20 * m.DL ) AS D20,	sum( d1.D21 * m.DL ) AS D21,	sum( d1.D22 * m.DL ) AS D22,	sum( d1.D23 * m.DL ) AS D23,	sum( d1.D24 * m.DL ) AS D24,	sum( d1.D25 * m.DL ) AS D25,	sum( d1.D26 * m.DL ) AS D26,	sum( d1.D27 * m.DL ) AS D27,	sum( d1.D28 * m.DL ) AS D28,	sum( d1.D29 * m.DL ) AS D29,	sum( d1.D30 * m.DL ) AS D30,	sum( d1.D31 * m.DL ) AS D31,	sum( d1.D32 * m.DL ) AS D32,	sum( d1.D33 * m.DL ) AS D33,	sum( d1.D34 * m.DL ) AS D34,	sum( d1.D35 * m.DL ) AS D35 FROM	BMTEMP D1	LEFT JOIN BasicMaterials M ON m.CPBH = d1.CPBM 	 GROUP BY	d1.XH,	d1.KHBM,	d1.CPBM,	D1.CPMC,	d1.DDYLH,	d1.SHSLH,	d1.kc,	D1.YYS,	D1.dys,	d1.WQWZH,	d1.PDSFCC,	D1.QUY,	d1.JHTS,	M.WLDL,	M.WLZL,	M.WLBH,	M.WLMC,	m.WLDW";
            dtrr = Tools.Common.SQLiteHelper.ExecuteDatatable(sql);
            return dtrr;
        }
        private DataTable n4()
        {
            DataTable dtrr = new DataTable();
            string sql = "SELECT D1.QUY AS 区分1,	M.WLDL AS 物料大类,	M.WLZL AS 物料子类,	M.WLBH AS 物料编码,	M.WLMC AS 物料名称,	m.WLDW as 物料单位,	sum( d1.D1 * m.DL ) AS D1,	sum( d1.D2 * m.DL ) AS D2,	sum( d1.D3 * m.DL ) AS D3,	sum( d1.D4 * m.DL ) AS D4,	sum( d1.D5 * m.DL ) AS D5,	sum( d1.D6 * m.DL ) AS D6,	sum( d1.D7 * m.DL ) AS D7,	sum( d1.D8 * m.DL ) AS D8,	sum( d1.D9 * m.DL ) AS D9,	sum( d1.D10 * m.DL ) AS D10,	sum( d1.D11 * m.DL ) AS D11,	sum( d1.D12 * m.DL ) AS D12,	sum( d1.D13 * m.DL ) AS D13,	sum( d1.D14 * m.DL ) AS D14,	sum( d1.D15 * m.DL ) AS D15,	sum( d1.D16 * m.DL ) AS D16,	sum( d1.D17 * m.DL ) AS D17,	sum( d1.D18 * m.DL ) AS D18,	sum( d1.D19 * m.DL ) AS D19,	sum( d1.D20 * m.DL ) AS D20,	sum( d1.D21 * m.DL ) AS D21,	sum( d1.D22 * m.DL ) AS D22,	sum( d1.D23 * m.DL ) AS D23,	sum( d1.D24 * m.DL ) AS D24,	sum( d1.D25 * m.DL ) AS D25,	sum( d1.D26 * m.DL ) AS D26,	sum( d1.D27 * m.DL ) AS D27,	sum( d1.D28 * m.DL ) AS D28,	sum( d1.D29 * m.DL ) AS D29,	sum( d1.D30 * m.DL ) AS D30,	sum( d1.D31 * m.DL ) AS D31,	sum( d1.D32 * m.DL ) AS D32,	sum( d1.D33 * m.DL ) AS D33,	sum( d1.D34 * m.DL ) AS D34,	sum( d1.D35 * m.DL ) AS D35 FROM	BMTEMP D1	LEFT JOIN BasicMaterials M ON m.CPBH = d1.CPBM 	where d1.quy='订单需求' GROUP BY	D1.QUY,	M.WLDL,	M.WLZL,	M.WLBH,	M.WLMC,	m.WLDW";
            DataTable dtrr1 = Tools.Common.SQLiteHelper.ExecuteDatatable(sql);
            //MessageBox.Show("dt1="+ dtrr1.Rows.Count);
            string sql1 = "SELECT D1.QUY AS 区分1,	M.WLDL AS 物料大类,	M.WLZL AS 物料子类,	M.WLBH AS 物料编码,	M.WLMC AS 物料名称,	m.WLDW as 物料单位,	sum( d1.D1 * m.DL ) AS D1,	sum( d1.D2 * m.DL ) AS D2,	sum( d1.D3 * m.DL ) AS D3,	sum( d1.D4 * m.DL ) AS D4,	sum( d1.D5 * m.DL ) AS D5,	sum( d1.D6 * m.DL ) AS D6,	sum( d1.D7 * m.DL ) AS D7,	sum( d1.D8 * m.DL ) AS D8,	sum( d1.D9 * m.DL ) AS D9,	sum( d1.D10 * m.DL ) AS D10,	sum( d1.D11 * m.DL ) AS D11,	sum( d1.D12 * m.DL ) AS D12,	sum( d1.D13 * m.DL ) AS D13,	sum( d1.D14 * m.DL ) AS D14,	sum( d1.D15 * m.DL ) AS D15,	sum( d1.D16 * m.DL ) AS D16,	sum( d1.D17 * m.DL ) AS D17,	sum( d1.D18 * m.DL ) AS D18,	sum( d1.D19 * m.DL ) AS D19,	sum( d1.D20 * m.DL ) AS D20,	sum( d1.D21 * m.DL ) AS D21,	sum( d1.D22 * m.DL ) AS D22,	sum( d1.D23 * m.DL ) AS D23,	sum( d1.D24 * m.DL ) AS D24,	sum( d1.D25 * m.DL ) AS D25,	sum( d1.D26 * m.DL ) AS D26,	sum( d1.D27 * m.DL ) AS D27,	sum( d1.D28 * m.DL ) AS D28,	sum( d1.D29 * m.DL ) AS D29,	sum( d1.D30 * m.DL ) AS D30,	sum( d1.D31 * m.DL ) AS D31,	sum( d1.D32 * m.DL ) AS D32,	sum( d1.D33 * m.DL ) AS D33,	sum( d1.D34 * m.DL ) AS D34,	sum( d1.D35 * m.DL ) AS D35 FROM	BMTEMP D1	LEFT JOIN BasicMaterials M ON m.CPBH = d1.CPBM 	where d1.quy='预测需求' GROUP BY	D1.QUY,	M.WLDL,	M.WLZL,	M.WLBH,	M.WLMC,	m.WLDW";
            DataTable dtrr2 = Tools.Common.SQLiteHelper.ExecuteDatatable(sql1);
            //MessageBox.Show("dt2=" + dtrr2.Rows.Count);
            dtrr = adddatatable(dtrr1, dtrr2);
            for (int i = 0; i < dtrr.Rows.Count; i++)
            {
                if (dtrr.Rows[i]["物料大类"].ToString() == "")
                {
                    dtrr.Rows.RemoveAt(i);
                }
            }
            //MessageBox.Show("dt3=" + dtrr.Rows.Count);
            index2 = dtrr.Rows.Count/2;
            return dtrr;
        }
        private void nCalcProductProcessNumber(DataTable dt1)
        {
            string connStr1 = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
            DataTable d1 = new DataTable();
            //清空数据表
            string sql = "delete from BPTEMP";
            Tools.Common.SQLiteHelper.ExecuteNonQuery(sql);
            using (SQLiteConnection con = new SQLiteConnection(connStr1))
            {
                con.Open();
                DbTransaction trans = con.BeginTransaction();//开始事务  
                SQLiteCommand cmd = new SQLiteCommand(con);

                cmd.CommandText = "INSERT INTO BPTEMP (XH,KHBM,CPBM,CPMC,DDYLH,SHSLH,KC,YYS,DYS,WQWZH,PDSFCC,QUY,JHTS,D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15,D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30,D31,D32,D33,D34,D35) VALUES(@XH,@KHBM,@CPBM,@CPMC,@DDYLH,@SHSLH,@KC,@YYS,@DYS,@WQWZH,@PDSFCC,@QUY,@JHTS,@D1,@D2,@D3,@D4,@D5,@D6,@D7,@D8,@D9,@D10,@D11,@D12,@D13,@D14,@D15,@D16,@D17,@D18,@D19,@D20,@D21,@D22,@D23,@D24,@D25,@D26,@D27,@D28,@D29,@D30,@D31,@D32,@D33,@D34,@D35)";
                foreach (DataRow dw in dt1.Rows)
                {
                    cmd.Parameters.Add(new SQLiteParameter("@XH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@KHBM", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@CPBM", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@DDYLH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@SHSLH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@KC", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@YYS", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@DYS", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@WQWZH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@PDSFCC", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@QUY", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@JHTS", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D1", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D2", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D3", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D4", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D5", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D6", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D7", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D8", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D9", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D10", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D11", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D12", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D13", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D14", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D15", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D16", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D17", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D18", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D19", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D20", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D21", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D22", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D23", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D24", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D25", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D26", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D27", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D28", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D29", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D30", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D31", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D32", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D33", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D34", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D35", DbType.String));
                    cmd.Parameters["@XH"].Value = ConvertStr(dw[0].ToString());
                    cmd.Parameters["@KHBM"].Value = ConvertStr(dw[1].ToString());
                    cmd.Parameters["@CPBM"].Value = ConvertStr(dw[2].ToString());
                    cmd.Parameters["@CPMC"].Value = ConvertStr(dw[3].ToString());
                    cmd.Parameters["@DDYLH"].Value = ConvertStr(dw[4].ToString());
                    cmd.Parameters["@SHSLH"].Value = ConvertStr(dw[5].ToString());
                    cmd.Parameters["@KC"].Value = ConvertStr(dw[6].ToString());
                    cmd.Parameters["@YYS"].Value = ConvertStr(dw[7].ToString());
                    cmd.Parameters["@DYS"].Value = ConvertStr(dw[8].ToString());
                    cmd.Parameters["@WQWZH"].Value = ConvertStr(dw[9].ToString());
                    cmd.Parameters["@PDSFCC"].Value = ConvertStr(dw[10].ToString());
                    cmd.Parameters["@QUY"].Value = ConvertStr(dw[11].ToString());
                    cmd.Parameters["@JHTS"].Value = ConvertStr(dw[12].ToString());
                    cmd.Parameters["@D1"].Value = ConvertStr(dw[13].ToString());
                    cmd.Parameters["@D2"].Value = ConvertStr(dw[14].ToString());
                    cmd.Parameters["@D3"].Value = ConvertStr(dw[15].ToString());
                    cmd.Parameters["@D4"].Value = ConvertStr(dw[16].ToString());
                    cmd.Parameters["@D5"].Value = ConvertStr(dw[17].ToString());
                    cmd.Parameters["@D6"].Value = ConvertStr(dw[18].ToString());
                    cmd.Parameters["@D7"].Value = ConvertStr(dw[19].ToString());
                    cmd.Parameters["@D8"].Value = ConvertStr(dw[20].ToString());
                    cmd.Parameters["@D9"].Value = ConvertStr(dw[21].ToString());
                    cmd.Parameters["@D10"].Value = ConvertStr(dw[22].ToString());
                    cmd.Parameters["@D11"].Value = ConvertStr(dw[23].ToString());
                    cmd.Parameters["@D12"].Value = ConvertStr(dw[24].ToString());
                    cmd.Parameters["@D13"].Value = ConvertStr(dw[25].ToString());
                    cmd.Parameters["@D14"].Value = ConvertStr(dw[26].ToString());
                    cmd.Parameters["@D15"].Value = ConvertStr(dw[27].ToString());
                    cmd.Parameters["@D16"].Value = ConvertStr(dw[28].ToString());
                    cmd.Parameters["@D17"].Value = ConvertStr(dw[29].ToString());
                    cmd.Parameters["@D18"].Value = ConvertStr(dw[30].ToString());
                    cmd.Parameters["@D19"].Value = ConvertStr(dw[31].ToString());
                    cmd.Parameters["@D20"].Value = ConvertStr(dw[32].ToString());
                    cmd.Parameters["@D21"].Value = ConvertStr(dw[33].ToString());
                    cmd.Parameters["@D22"].Value = ConvertStr(dw[34].ToString());
                    cmd.Parameters["@D23"].Value = ConvertStr(dw[35].ToString());
                    cmd.Parameters["@D24"].Value = ConvertStr(dw[36].ToString());
                    cmd.Parameters["@D25"].Value = ConvertStr(dw[37].ToString());
                    cmd.Parameters["@D26"].Value = ConvertStr(dw[38].ToString());
                    cmd.Parameters["@D27"].Value = ConvertStr(dw[39].ToString());
                    cmd.Parameters["@D28"].Value = ConvertStr(dw[40].ToString());
                    try
                    {
                        cmd.Parameters["@D29"].Value = ConvertStr(dw[41].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D29"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D30"].Value = ConvertStr(dw[42].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D30"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D31"].Value = ConvertStr(dw[43].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D31"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D32"].Value = ConvertStr(dw[44].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D32"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D33"].Value = ConvertStr(dw[45].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D33"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D34"].Value = ConvertStr(dw[46].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D34"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D35"].Value = ConvertStr(dw[47].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D35"].Value = "";
                    }
                    cmd.ExecuteNonQuery();
                }
                trans.Commit();//事务提交   
            }
        }

        private void nCalcMaterialNumber(DataTable dt1)
        {
            string connStr1 = @"Data Source=" + "BasicData.db;Initial Catalog=sqlite;Integrated Security=True;Max Pool Size=10";
            DataTable d1 = new DataTable();
            //清空数据表
            string sql = "delete from BMTEMP";
            Tools.Common.SQLiteHelper.ExecuteNonQuery(sql);
            using (SQLiteConnection con = new SQLiteConnection(connStr1))
            {
                con.Open();
                DbTransaction trans = con.BeginTransaction();//开始事务  
                SQLiteCommand cmd = new SQLiteCommand(con);

                cmd.CommandText = "INSERT INTO BMTEMP (XH,KHBM,CPBM,CPMC,DDYLH,SHSLH,KC,YYS,DYS,WQWZH,PDSFCC,QUY,JHTS,D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15,D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30,D31,D32,D33,D34,D35) VALUES(@XH,@KHBM,@CPBM,@CPMC,@DDYLH,@SHSLH,@KC,@YYS,@DYS,@WQWZH,@PDSFCC,@QUY,@JHTS,@D1,@D2,@D3,@D4,@D5,@D6,@D7,@D8,@D9,@D10,@D11,@D12,@D13,@D14,@D15,@D16,@D17,@D18,@D19,@D20,@D21,@D22,@D23,@D24,@D25,@D26,@D27,@D28,@D29,@D30,@D31,@D32,@D33,@D34,@D35)";
                foreach (DataRow dw in dt1.Rows)
                {
                    cmd.Parameters.Add(new SQLiteParameter("@XH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@KHBM", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@CPBM", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@CPMC", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@DDYLH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@SHSLH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@KC", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@YYS", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@DYS", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@WQWZH", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@PDSFCC", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@QUY", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@JHTS", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D1", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D2", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D3", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D4", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D5", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D6", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D7", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D8", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D9", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D10", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D11", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D12", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D13", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D14", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D15", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D16", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D17", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D18", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D19", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D20", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D21", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D22", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D23", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D24", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D25", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D26", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D27", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D28", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D29", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D30", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D31", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D32", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D33", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D34", DbType.String));
                    cmd.Parameters.Add(new SQLiteParameter("@D35", DbType.String));
                    cmd.Parameters["@XH"].Value = ConvertStr(dw[0].ToString());
                    cmd.Parameters["@KHBM"].Value = ConvertStr(dw[1].ToString());
                    cmd.Parameters["@CPBM"].Value = ConvertStr(dw[2].ToString());
                    cmd.Parameters["@CPMC"].Value = ConvertStr(dw[3].ToString());
                    cmd.Parameters["@DDYLH"].Value = ConvertStr(dw[4].ToString());
                    cmd.Parameters["@SHSLH"].Value = ConvertStr(dw[5].ToString());
                    cmd.Parameters["@KC"].Value = ConvertStr(dw[6].ToString());
                    cmd.Parameters["@YYS"].Value = ConvertStr(dw[7].ToString());
                    cmd.Parameters["@DYS"].Value = ConvertStr(dw[8].ToString());
                    cmd.Parameters["@WQWZH"].Value = ConvertStr(dw[9].ToString());
                    cmd.Parameters["@PDSFCC"].Value = ConvertStr(dw[10].ToString());
                    cmd.Parameters["@QUY"].Value = ConvertStr(dw[11].ToString());
                    cmd.Parameters["@JHTS"].Value = ConvertStr(dw[12].ToString());
                    cmd.Parameters["@D1"].Value = ConvertStr(dw[13].ToString());
                    cmd.Parameters["@D2"].Value = ConvertStr(dw[14].ToString());
                    cmd.Parameters["@D3"].Value = ConvertStr(dw[15].ToString());
                    cmd.Parameters["@D4"].Value = ConvertStr(dw[16].ToString());
                    cmd.Parameters["@D5"].Value = ConvertStr(dw[17].ToString());
                    cmd.Parameters["@D6"].Value = ConvertStr(dw[18].ToString());
                    cmd.Parameters["@D7"].Value = ConvertStr(dw[19].ToString());
                    cmd.Parameters["@D8"].Value = ConvertStr(dw[20].ToString());
                    cmd.Parameters["@D9"].Value = ConvertStr(dw[21].ToString());
                    cmd.Parameters["@D10"].Value = ConvertStr(dw[22].ToString());
                    cmd.Parameters["@D11"].Value = ConvertStr(dw[23].ToString());
                    cmd.Parameters["@D12"].Value = ConvertStr(dw[24].ToString());
                    cmd.Parameters["@D13"].Value = ConvertStr(dw[25].ToString());
                    cmd.Parameters["@D14"].Value = ConvertStr(dw[26].ToString());
                    cmd.Parameters["@D15"].Value = ConvertStr(dw[27].ToString());
                    cmd.Parameters["@D16"].Value = ConvertStr(dw[28].ToString());
                    cmd.Parameters["@D17"].Value = ConvertStr(dw[29].ToString());
                    cmd.Parameters["@D18"].Value = ConvertStr(dw[30].ToString());
                    cmd.Parameters["@D19"].Value = ConvertStr(dw[31].ToString());
                    cmd.Parameters["@D20"].Value = ConvertStr(dw[32].ToString());
                    cmd.Parameters["@D21"].Value = ConvertStr(dw[33].ToString());
                    cmd.Parameters["@D22"].Value = ConvertStr(dw[34].ToString());
                    cmd.Parameters["@D23"].Value = ConvertStr(dw[35].ToString());
                    cmd.Parameters["@D24"].Value = ConvertStr(dw[36].ToString());
                    cmd.Parameters["@D25"].Value = ConvertStr(dw[37].ToString());
                    cmd.Parameters["@D26"].Value = ConvertStr(dw[38].ToString());
                    cmd.Parameters["@D27"].Value = ConvertStr(dw[39].ToString());
                    cmd.Parameters["@D28"].Value = ConvertStr(dw[40].ToString());
                    try
                    {
                        cmd.Parameters["@D29"].Value = ConvertStr(dw[41].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D29"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D30"].Value = ConvertStr(dw[42].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D30"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D31"].Value = ConvertStr(dw[43].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D31"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D32"].Value = ConvertStr(dw[44].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D32"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D33"].Value = ConvertStr(dw[45].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D33"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D34"].Value = ConvertStr(dw[46].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D34"].Value = "";
                    }
                    try
                    {
                        cmd.Parameters["@D35"].Value = ConvertStr(dw[47].ToString());
                    }
                    catch
                    {
                        cmd.Parameters["@D35"].Value = "";
                    }
                    cmd.ExecuteNonQuery();
                }
                trans.Commit();//事务提交   
            }
        }
    }

    #endregion

}

