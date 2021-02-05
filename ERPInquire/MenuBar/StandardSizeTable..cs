using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERPInquire.CustomControl
{
    public partial class StandardSizeTable : UserControl
    {
        DataTable dt = new DataTable();
        public StandardSizeTable()
        {
            InitializeComponent();
        }
        #region  按钮事件
        //查询
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            string sql = "select Name as 开数,IsDegress AS '正度(cm×cm)',BigDegress AS '大度(cm×cm)'  from StandardSizeTable  where Name like '%" + toolStripTextBox1.Text + "%'   order by convert(int,ID)";
            dt = SqlHelper.ExecuteDataTable(sql);
            SqlHelper.GetConnection().Close();
            dataGridView1.DataSource = dt;
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Azure;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
            toolStripLabel3.Text = "总行数:" + (dt.Rows.Count).ToString();
            dataGridView1.RowsDefaultCellStyle.Font = new Font("微软雅黑", 8, FontStyle.Regular);
            dataGridView1.Refresh();
        }

        //导出EXCEL
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MyMIS.ExcleIO.OutExcel("制版纸张规格尺寸表", dt);
        }
        #endregion

    }
}
