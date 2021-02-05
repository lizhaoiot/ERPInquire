using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace Excel
{
    public class ExcelHelper
    {
        /// <summary>
        /// 类版本
        /// </summary>
        public string version
        {
            get { return "0.1"; }
        }
        readonly int EXCEL03_MaxRow = 65535;

        /// <summary>
        /// 将DataTable转换为excel2003格式。
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        //public byte[] DataTable2Excel(DataTable dt, string sheetName)
        //{

        //    IWorkbook book = new HSSFWorkbook();
        //    if (dt.Rows.Count < EXCEL03_MaxRow)
        //        DataWrite2Sheet(dt, 0, dt.Rows.Count - 1, book, sheetName);
        //    else
        //    {
        //        int page = dt.Rows.Count / EXCEL03_MaxRow;
        //        for (int i = 0; i < page; i++)
        //        {
        //            int start = i * EXCEL03_MaxRow;
        //            int end = (i * EXCEL03_MaxRow) + EXCEL03_MaxRow - 1;
        //            DataWrite2Sheet(dt, start, end, book, sheetName + i.ToString());
        //        }
        //        int lastPageItemCount = dt.Rows.Count % EXCEL03_MaxRow;
        //        DataWrite2Sheet(dt, dt.Rows.Count - lastPageItemCount, lastPageItemCount, book, sheetName + page.ToString());
        //    }
        //    MemoryStream ms = new MemoryStream();
        //    book.Write(ms);
        //    return ms.ToArray();
        //}
        //private void DataWrite2Sheet(DataTable dt, int startRow, int endRow, IWorkbook book, string sheetName)
        //{
        //    ISheet sheet = book.CreateSheet(sheetName);
        //    IRow header = sheet.CreateRow(0);
        //    for (int i = 0; i < dt.Columns.Count; i++)
        //    {
        //        ICell cell = header.CreateCell(i);
        //        string val = dt.Columns[i].Caption ?? dt.Columns[i].ColumnName;
        //        cell.SetCellValue(val);
        //    }
        //    int rowIndex = 1;
        //    for (int i = startRow; i <= endRow; i++)
        //    {
        //        DataRow dtRow = dt.Rows[i];
        //        IRow excelRow = sheet.CreateRow(rowIndex++);
        //        for (int j = 0; j < dtRow.ItemArray.Length; j++)
        //        {
        //            excelRow.CreateCell(j).SetCellValue(dtRow[j].ToString());
        //        }
        //    }

        //}
    }
}