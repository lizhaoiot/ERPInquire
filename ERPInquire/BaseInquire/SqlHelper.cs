using System;
using System.Configuration;
using System.ComponentModel;
using System.Collections;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Win32;
using System.IO;
using System.Data.Common;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using Com.Hui.iMRP.Utils;


public class SqlHelper
{
    public static DataTable GetDgvToTable(DataGridView dgv)
    {
        DataTable dt = new DataTable();
        // 列强制转换
        for (int count = 0; count < dgv.Columns.Count; count++)
        {
            DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
            dt.Columns.Add(dc);
        }
        // 循环行
        for (int count = 0; count < dgv.Rows.Count; count++)
        {
            DataRow dr = dt.NewRow();
            for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
            {
                dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
            }
            dt.Rows.Add(dr);
        }
        return dt;
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
                if (dt.Columns[i].ColumnName.Contains("日期") || dt.Columns[i].ColumnName.Contains("量") || dt.Columns[i].ColumnName.Contains("数") || dt.Columns[i].ColumnName.Contains("长") || dt.Columns[i].ColumnName.Contains("宽") || dt.Columns[i].ColumnName.Contains("高") || dt.Columns[i].ColumnName.Contains("面积") || dt.Columns[i].ColumnName.Contains("重") || dt.Columns[i].ColumnName.Contains("度"))
                    continue;
                ls.Add(dt.Columns[i].ColumnName);
            }
        }
        return ls;
    }
    public static DataTable GetUserList(String connectionStr, String sql, SqlParameter[] prams)
    {
        DataTable table = null;
        {
            SqlConnection specialCon = new SqlConnection(connectionStr);
            specialCon.Open();
            using (SqlCommand cmd = new SqlCommand(sql, specialCon))
            {
                cmd.Parameters.AddRange(prams);
                DataSet dataSet = new DataSet();
                System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = cmd;
                adapter.Fill(dataSet);
                table = dataSet.Tables[0];
            }
            specialCon.Close();
        }

        return table;
    }
    public static SqlConnection GetConnection()
    {
        SqlConnect.ConnectStr connect = new SqlConnect.ConnectStr("ERP");
        string s = connect.GetConnectStr("ERP");
        return GetConnection(s);
        //return GetConnection("packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy");
    }
    public static SqlConnection GetConnection(String connstr)
    {
        SqlConnection conn = null;
        {
            conn = new SqlConnection(connstr);
            conn.Open();
        }
        return conn;
    }
    public static void BulkCopy(DataTable src, String dest)
    {
        using (SqlConnection conn = GetConnection())
        {
            var sbc = new SqlBulkCopy(conn);
            sbc.DestinationTableName = dest;
            sbc.WriteToServer(src);
        }
    }
    public static void BulkCopy(DataTable src, String dest, SqlConnection conn, SqlTransaction t)
    {
        var sbc = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, t);
        sbc.DestinationTableName = dest;
        sbc.WriteToServer(src);
    }
    /// <summary>
    /// 执行一条不需要返回值SQL命令，比如插入、删除
    /// </summary>
    /// <param name="commtext"></param>
    public static void ExecCommand(string commtext)
    {
        using (SqlConnection conn = GetConnection())
        {
            SqlCommand cmd = new SqlCommand(commtext, conn);
            cmd.CommandType = CommandType.Text;
            cmd.ExecuteNonQuery();
            cmd.Dispose();
        }
    }
    /// <summary>
    /// 执行多条不需要返回值SQL命令，比如插入、删除
    /// </summary>
    /// <param name="commtext">多条SQL语句</param>
    public static void ExecCommand(string[] commtext)
    {
        using (SqlConnection conn = GetConnection())
        {
            for (int i = 0; i < commtext.Length; i++)
            {
                SqlCommand cmd = new SqlCommand(commtext[i], conn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
        }
    }
    /// <summary>
    /// 执行一条不需要返回值SQL命令，比如插入、删除，可以带一个参数数组
    /// </summary>
    /// <param name="commtext"></param>
    /// <param name="prams"></param>
    public static int ExecCommand(string commtext, SqlParameter[] prams)
    {
        using (SqlConnection conn = GetConnection())
        {
            int ret = -1;
            SqlCommand cmd = new SqlCommand(commtext, conn);
            cmd.CommandType = CommandType.Text;

            if (prams != null)
            {
                cmd.Parameters.AddRange(prams);
            }
            ret = cmd.ExecuteNonQuery();
            cmd.Dispose();
            return ret;
        }
    }
    public static void ExecCommand(string commtext, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
    {
        SqlCommand cmd = new SqlCommand(commtext, conn);
        cmd.CommandType = CommandType.Text;
        cmd.Transaction = t;
        if (prams != null)
        {
            cmd.Parameters.AddRange(prams);
        }
        cmd.ExecuteNonQuery();
    }
    public static void ExecCommand(string commtext, SqlConnection conn, SqlTransaction t)
    {
        SqlCommand cmd = new SqlCommand(commtext, conn);
        cmd.CommandType = CommandType.Text;
        cmd.Transaction = t;
        cmd.ExecuteNonQuery();
    }
    public static int ExecStoredProcedure(String procedureName, SqlParameter[] prams)
    {
        using (SqlConnection conn = GetConnection())
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 600;
            cmd.Connection = conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procedureName;   //存储过程名称 
            if (prams != null)
            {
                cmd.Parameters.AddRange(prams);
            }
            int retValue = cmd.ExecuteNonQuery();
            cmd.Dispose();
            return retValue;
        }
    }
    public static DataTable ExecStoredProcedureDataTable(String procedureName, SqlParameter[] prams)
    {
        using (SqlConnection conn = GetConnection())
        {
            DataTable table = null;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 600;
            cmd.Connection = conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procedureName;   //存储过程名称 
            if (prams != null)
            {
                cmd.Parameters.AddRange(prams);
            }
            DataSet dataSet = new DataSet();
            System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = cmd;
            adapter.Fill(dataSet);
            if (dataSet.Tables.Count == 0)
            {
                return null;
            }
            table = dataSet.Tables[0];
            cmd.Dispose();
            return table;
        }
    }

    public static DataSet ExecStoredProcedureDataSet(String procedureName, SqlParameter[] prams)
    {
        using (SqlConnection conn = GetConnection())
        {
            DataSet dataSet = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 600;
            cmd.Connection = conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procedureName; //存储过程名称 
            if (prams != null)
            {
                cmd.Parameters.AddRange(prams);
            }
            System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = cmd;
            adapter.Fill(dataSet);
            cmd.Dispose();
            return dataSet;
        }
    }
    public static DataSet ExecStoredProcedureDataSet(String procedureName, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
    {
        //using (SqlConnection conn = GetConnection())
        {
            DataSet dataSet = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 600;
            cmd.Connection = conn;
            cmd.Transaction = t;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procedureName;   //存储过程名称 
            if (prams != null)
            {
                cmd.Parameters.AddRange(prams);
            }
            System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = cmd;
            adapter.Fill(dataSet);
            return dataSet;
        }
    }
    public static DataTable ExecStoredProcedureDataTable(String procedureName, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
    {
        DataTable table = null;
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 600;
        cmd.Connection = conn;
        cmd.Transaction = t;
        cmd.CommandType = CommandType.StoredProcedure;
        cmd.CommandText = procedureName;   //存储过程名称
        if (prams != null)
        {
            cmd.Parameters.AddRange(prams);
        }
        DataSet dataSet = new DataSet();
        System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
        adapter.SelectCommand = cmd;
        adapter.Fill(dataSet);
        table = dataSet.Tables[0];
        return table;
    }
    public static int ExecStoredProcedure(String procedureName, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
    {
        SqlCommand cmd = new SqlCommand();
        cmd.Connection = conn;
        cmd.Transaction = t;
        cmd.CommandType = CommandType.StoredProcedure;
        cmd.CommandText = procedureName;   //存储过程名称 
        if (prams != null)
        {
            cmd.Parameters.AddRange(prams);
        }
        int retValue = cmd.ExecuteNonQuery();
        return retValue;
    }
    /// <summary>
    /// 执行一条统计类的SQL语句，返回一个整数
    /// </summary>
    /// <param name="CommText"></param>
    /// <returns></returns>
    public static int ExecScalar(string CommText)
    {
        using (SqlConnection conn = GetConnection())
        {
            SqlCommand cmd = new SqlCommand(CommText, conn);
            cmd.CommandType = CommandType.Text;
            return cmd.ExecuteScalar() != DBNull.Value ? (int)cmd.ExecuteScalar() : 0;
        }
    }
    public static int ExecScalar(string CommText, SqlConnection TConn)
    {
        SqlCommand cmd = new SqlCommand(CommText, TConn);
        cmd.CommandType = CommandType.Text;
        return cmd.ExecuteScalar() != DBNull.Value ? (int)cmd.ExecuteScalar() : 0;
    }
    public static int ExecScalar(string CommText, SqlTransaction t)
    {
        using (SqlConnection conn = GetConnection())
        {
            SqlCommand cmd = new SqlCommand(CommText, conn);
            cmd.Transaction = t;
            cmd.CommandType = CommandType.Text;
            return cmd.ExecuteScalar() != DBNull.Value ? Convert.ToInt32(cmd.ExecuteScalar()) : 0;
        }
    }
    public static int ExecScalar(string CommText, SqlConnection TConn, SqlTransaction t)
    {
        SqlCommand cmd = new SqlCommand(CommText, TConn);
        cmd.Transaction = t;
        cmd.CommandType = CommandType.Text;
        return cmd.ExecuteScalar() != DBNull.Value ? Convert.ToInt32(cmd.ExecuteScalar()) : 0;
    }
    /// <summary>
    /// 执行一条统计类的SQL语句，返回一个整数。可以带参数
    /// </summary>
    /// <param name="CommText"></param>
    /// <param name="prams"></param>
    /// <returns></returns>
    public static int ExecScalar(string CommText, SqlParameter[] prams)
    {
        using (SqlConnection conn = GetConnection())
        {
            SqlCommand cmd = new SqlCommand(CommText, conn);
            cmd.CommandType = CommandType.Text;
            if (prams != null)
            {
                cmd.Parameters.AddRange(prams);
            }
            return cmd.ExecuteScalar() != DBNull.Value ? Convert.ToInt32(cmd.ExecuteScalar()) : 0;
        }
    }
    public static int ExecScalar(string CommText, SqlParameter[] prams, SqlConnection TConn)
    {
        SqlCommand cmd = new SqlCommand(CommText, TConn);
        cmd.CommandType = CommandType.Text;
        if (prams != null)
        {
            foreach (SqlParameter parameter in prams)
                cmd.Parameters.Add(parameter);
        }
        return cmd.ExecuteScalar() != DBNull.Value ? Convert.ToInt32(cmd.ExecuteScalar()) : 0;
    }
    /// <summary>
    /// 执行返回DataTable，不带参的查询语句
    /// </summary>
    /// <param name="sql"></param>
    /// <returns></returns>
    public static DataTable ExecuteDataTable(String sql)
    {
        using (SqlConnection conn = GetConnection())
        {
            DataTable table = null;
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                DataSet dataSet = new DataSet();
                System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = cmd;
                adapter.Fill(dataSet);
                table = dataSet.Tables[0];
                cmd.Parameters.Clear();
            }
            return table;
        }
    }
    public static DataTable ExecuteDataTable(String sql, SqlConnection conn, SqlTransaction t)
    {
        DataTable table = null;
        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            cmd.Transaction = t;
            DataSet dataSet = new DataSet();
            System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = cmd;
            adapter.Fill(dataSet);
            table = dataSet.Tables[0];
            cmd.Parameters.Clear();
        }
        return table;
    }
    /// <summary>
    /// 执行返回DataTable,带参的查询语句
    /// </summary>
    /// <param name="sql"></param>
    /// <returns></returns>
    public static DataTable ExecuteDataTable(String sql, DbParameter[] parameters)
    {
        DataTable table = null;
        using (SqlConnection conn = GetConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddRange(parameters);
                cmd.CommandTimeout = 600;
                DataSet dataSet = new DataSet();
                System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = cmd;
                adapter.Fill(dataSet);
                table = dataSet.Tables[0];
                cmd.Parameters.Clear();
            }//using
            return table;
        }
    }
    public static DataTable ExecuteDataTable(String sql, DbParameter[] parameters, String connectionString)
    {
        DataTable table = null;
        using (SqlConnection conn = GetConnection(connectionString))
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddRange(parameters);
                cmd.CommandTimeout = 600;
                DataSet dataSet = new DataSet();
                System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = cmd;
                adapter.Fill(dataSet);
                table = dataSet.Tables[0];
                cmd.Parameters.Clear();
            }//using
            return table;
        }
    }
    public static DataTable ExecuteDataTable(String sql, DbParameter[] parameters, SqlConnection conn, SqlTransaction t)
    {
        DataTable table = null;
        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            cmd.Transaction = t;
            cmd.Parameters.AddRange(parameters);
            DataSet dataSet = new DataSet();
            System.Data.Common.DbDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = cmd;
            adapter.Fill(dataSet);
            table = dataSet.Tables[0];
            cmd.Parameters.Clear();
        }//using
        return table;
    }
    /// <summary>
    /// 执行一条SQL语句，并返回一个命名了数据表的数据集
    /// </summary>
    /// <param name="cmdtext">SQL语句</param>
    /// <param name="tablename">表名</param>
    /// <param name="ds">数据集</param>
    public static void Exec4DS(string CmdText, string TableName, out DataSet DS)
    {
        using (SqlConnection conn = GetConnection())
        {
            DS = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(CmdText, conn);
            da.Fill(DS, TableName);
        }
    }
    public static DataSet ExecuteDataSet(String sql, SqlParameter[] prams)
    {
        using (SqlConnection conn = GetConnection())
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.CommandType = CommandType.Text;
            if (prams != null)
            {
                cmd.Parameters.AddRange(prams);
            }
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);
            cmd.Parameters.Clear();
            return ds;
        }
    }
    /// <summary>
    /// 一次执行多条SQL语句，将多个表填入一个DataSet
    /// </summary>
    /// <param name="sqlCmds">SQL语句数组</param>
    /// <param name="TableNames">表名数组</param>
    /// <param name="DS">数据集名</param>
    public static void ExecXDS(string[] sqlCmds, string[] TableNames, out DataSet DS)
    {
        using (SqlConnection conn = GetConnection())
        {
            DS = new DataSet();
            for (int i = 0; i < sqlCmds.Length; i++)
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlCmds[i], conn);
                da.Fill(DS, TableNames[i]);
            }
        }
    }
    public static void ExecXDS(string[] sqlCmds, string[] TableNames, out DataSet DS, String connectionString)
    {
        using (SqlConnection conn = GetConnection(connectionString))
        {
            DS = new DataSet();
            for (int i = 0; i < sqlCmds.Length; i++)
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlCmds[i], conn);
                da.Fill(DS, TableNames[i]);
            }
        }
    }
    /// <summary>
    /// 一条语句允许有一个参数且顺序要对应
    /// </summary>
    /// <param name="sqlCmds"></param>
    /// <param name="TableNames"></param>
    /// <param name="prams"></param>
    /// <param name="DS"></param>
    public static void ExecXDS(string[] sqlCmds, string[] TableNames, SqlParameter[] prams, out DataSet DS)
    {
        using (SqlConnection conn = GetConnection())
        {
            DS = new DataSet();
            for (int i = 0; i < sqlCmds.Length; i++)
            {
                SqlCommand cmd = new SqlCommand(sqlCmds[i]);
                cmd.Parameters.Add(prams[i]);
                cmd.Connection = conn;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(DS, TableNames[i]);
            }
        }
    }
    /// <summary>
    /// 一条语句可以传多个参数
    /// </summary>
    /// <param name="sqlCmds"></param>
    /// <param name="TableNames"></param>
    /// <param name="prams"></param>
    /// <param name="DS"></param>
    public static void ExecXDS(string[] sqlCmds, string[] TableNames, List<SqlParameter[]> prams, out DataSet DS)
    {
        using (SqlConnection conn = GetConnection())
        {
            DS = new DataSet();
            for (int i = 0; i < sqlCmds.Length; i++)
            {
                SqlCommand cmd = new SqlCommand(sqlCmds[i]);
                //foreach(SqlParameter[] pramArr in prams){
                SqlParameter[] pramArr = prams[i];//按顺序索引各自的参数数组
                                                  // for (int j = 0; j < pramArr.Length; j++)
                cmd.Parameters.AddRange(pramArr);
                cmd.Connection = conn;
                //}
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(DS, TableNames[i]);
            }
        }
    }
    public static DateTime GetDateTimeFromSQL(SqlConnection conn, SqlTransaction t)
    {
        string sql = "select getdate()";
        SqlCommand cmd = new SqlCommand(sql, conn);
        cmd.Transaction = t;
        DateTime dt;
        dt = (DateTime)cmd.ExecuteScalar();
        return dt;

    }
    public static DateTime GetDateTimeFromSQL()
    {
        using (SqlConnection conn = GetConnection())
        {
            string sql = "select getdate()";
            SqlCommand cmd = new SqlCommand(sql, conn);
            DateTime dt;
            dt = (DateTime)cmd.ExecuteScalar();
            return dt;
        }
    }

    public static void BulkToDB(DataTable dt,string Tempname,string[] str)
    {
        SqlConnection sqlConn = GetConnection("packet size=4096;user id=sa;pwd=;data source=192.168.0.97;persist security info=False;initial catalog=hy");
        SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn);
        bulkCopy.BulkCopyTimeout =60*60*100;
        bulkCopy.DestinationTableName = Tempname;
        for (int i = 0; i < str.Length; i++)
        {
          bulkCopy.ColumnMappings.Add(i, str[i]);
        }
        bulkCopy.BatchSize = dt.Rows.Count;
        try
        {
            if (dt != null && dt.Rows.Count != 0)
            {
                bulkCopy.WriteToServer(dt);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
        finally
        {
            sqlConn.Close();
            if (bulkCopy != null)
            {
                bulkCopy.Close();
            }
        }
    }
}