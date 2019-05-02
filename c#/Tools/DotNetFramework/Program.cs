using DotNetFramework.NPOI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetFramework
{
    class Program
    {
        static void Main(string[] args)
        {

            //NpoiUtil.ModifyExcelDemo(@"E:\test.xlsx");

            //==============================================================//
            //NpoiUtil.ExcelToDataTable(@"E:\test.xlsx");


            //==============================================================//
            //DataTable dt = new DataTable();
            //DataColumn column = new DataColumn("a");
            //dt.Columns.Add(column);
            //column = new DataColumn("b");
            //dt.Columns.Add(column);

            //dt.TableName = "999";
            //var row = dt.NewRow();
            //row[0] = "11";
            //row[1] = "22";
            //dt.Rows.Add(row);
            //var row1 = dt.NewRow();
            //row1[0] = "11";
            //row1[1] = "22";
            //dt.Rows.Add(row1);
            //NpoiUtil.CreateExcelByDataTable(dt, @"E:\FieldedAddresses.xlsx");
        }
    }
}
