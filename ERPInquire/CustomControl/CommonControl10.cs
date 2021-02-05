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

namespace ERPInquire.CustomControl
{
    public partial class CommonControl10 : UserControl
    {
        #region 定量

        public string NodeName;
        public string TimeKey;
        public string ProName;
        #endregion

        #region 变量

        DateTimePicker sdate;
        DateTimePicker edate;
        TextBox filter;
        public string sqlstr;
        public ConnectStr connect;
        DataTable dt = new DataTable();
        DateTime starttime;
        DateTime endtime;

        #endregion

        #region 初始化
        public CommonControl10()
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
            starttime = this.sdate.Value;
            endtime = this.edate.Value;
            MyMIS.AddToolstrip.AddDTPtoToolstrip(4, sdate, toolStrip1);
            MyMIS.AddToolstrip.AddDTPtoToolstrip(6, edate, toolStrip1);
        }
        private void CommonControl10_Load(object sender, EventArgs e)
        {
            toolStripLabel2.Text = TimeKey;
            toolStripComboBox1.SelectedIndex = 0;
        }
        #endregion

        #region 窗体事件
        //查询
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//等待
            if (toolStripTextBox1.Text.Equals(""))
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
                        para[0].Value = toolStripComboBox1.Text;
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
                            List<string> ls = SqlHelper.GetColumnsByDataTable(dt);
                            toolStripComboBox2.ComboBox.DataSource = ls;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
            else
            {
                try
              {
                //获得现在所选列的列索引
                int i = toolStripComboBox2.ComboBox.SelectedIndex;
                string dtcolounname = toolStripComboBox2.SelectedItem.ToString();
                DataRow[] dr = dt.Select(dtcolounname + " like '%" + toolStripTextBox1.Text + "%'");
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
        //导出EXCEL
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.Export2Excel(NodeName, SqlHelper.GetDgvToTable(dataGridView1));
        }
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < e.RowCount; i++)
                this.dataGridView1.Rows[e.RowIndex + i].HeaderCell.Value = (e.RowIndex + i + 1).ToString();
        }

        #endregion

        #region 方法

        #endregion

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(toolStripTextBox1.Text))
            {
                toolStripComboBox1.Enabled = true;
            }
            else
            {
                toolStripComboBox1.Enabled = false;
            }
        }
        //计划部专用
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            DataTable dcf = SqlHelper.GetDgvToTable(dataGridView1);

            DataTable tblDatas = new DataTable();
            DataColumn dc = null;

            dc = tblDatas.Columns.Add("印刷机台", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("生产施工单单号", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("产品编号", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("产品名称", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("部件名称", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("应产数", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("制版方式", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("上机尺寸", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("正面色数", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("反面色数", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("物料名称", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("工序描述", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("正面颜色", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("反面颜色", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("工序名称", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("合大张数", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("产品类别", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("产品构成", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("成品尺寸", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("刀版编号", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("物料编号", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("总P数", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("模帖P数", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("生产数量", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("印刷数量", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("", Type.GetType("System.String"));

            for (int i = 0; i < dcf.Rows.Count; i++)
            {
                DataRow newRow;
                newRow = tblDatas.NewRow();
                newRow[0] = dcf.Rows[i]["印刷机台"].ToString();
                newRow[1] = dcf.Rows[i]["生产施工单单号"].ToString();
                newRow[2] = dcf.Rows[i]["产品编号"].ToString();
                newRow[3] = dcf.Rows[i]["产品名称"].ToString();
                newRow[4] = dcf.Rows[i]["部件名称"].ToString();
                newRow[5] = dcf.Rows[i]["应产数"].ToString();
                newRow[6] = dcf.Rows[i]["制版方式"].ToString();
                newRow[7] = dcf.Rows[i]["上机尺寸"].ToString();
                newRow[8] = dcf.Rows[i]["正面色数"].ToString();
                newRow[9] = "";
                newRow[10] = dcf.Rows[i]["反面色数"].ToString();
                newRow[11] = "";
                newRow[12] = "";
                newRow[13] = "";
                newRow[14] = "";
                newRow[15] = "";
                newRow[16] = "";
                newRow[17] = "";
                newRow[18] = "";
                newRow[19] = dcf.Rows[i]["物料名称"].ToString();
                newRow[20] = "";
                newRow[21] = "";
                newRow[22] = dcf.Rows[i]["工序描述"].ToString();
                newRow[23] = dcf.Rows[i]["正面颜色"].ToString();
                newRow[24] = dcf.Rows[i]["反面颜色"].ToString();
                newRow[25] = dcf.Rows[i]["工序名称"].ToString();
                newRow[26] = "";
                newRow[27] = "";
                newRow[28] = "";
                newRow[29] = "";
                newRow[30] = "";
                newRow[31] = dcf.Rows[i]["合大张数"].ToString();
                newRow[32] = "";
                newRow[33] = "";
                newRow[34] = dcf.Rows[i]["产品类别"].ToString();
                newRow[35] = dcf.Rows[i]["产品构成"].ToString();
                newRow[36] = dcf.Rows[i]["成品尺寸"].ToString();
                newRow[37] = dcf.Rows[i]["刀版编号"].ToString();
                newRow[38] = dcf.Rows[i]["物料编号"].ToString();
                newRow[39] = dcf.Rows[i]["总P数"].ToString();
                newRow[40] = dcf.Rows[i]["模帖P数"].ToString();
                newRow[41] = dcf.Rows[i]["生产数量"].ToString();
                newRow[42] = dcf.Rows[i]["印刷数量"].ToString();
                newRow[43] = "";
                newRow[44] = "";
                newRow[45] = "";
                tblDatas.Rows.Add(newRow);
            }
            MyMIS.ExcleIO.Export2Excel(NodeName, tblDatas);
        }
    }
}
