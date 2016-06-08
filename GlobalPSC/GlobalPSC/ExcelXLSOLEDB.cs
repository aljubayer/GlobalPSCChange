using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ManiacProject.Libs
{
    public class ExecuteNonQueryOnExcelXLS
    {
        System.Data.OleDb.OleDbConnection MyConnection;
        System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

        public ExecuteNonQueryOnExcelXLS(string dbFileName)
        {
            MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + dbFileName + "';Extended Properties=Excel 8.0;");
            MyConnection.Open();
        }
        public int ExecuteCommandOnExcelFile(string nonQuery)
        {
            try
            {
                myCommand.Connection = MyConnection;
                myCommand.CommandText = nonQuery;
                int affectedRows = myCommand.ExecuteNonQuery();
                return affectedRows;
            }
            catch (Exception exception)
            {
                MyConnection.Close();
                throw new Exception(exception.Message);
            }

        }

        public static DataSet ReadFromExcelFile(string file, string sheetName)
        {
            string query = "select * from [" + sheetName + "$]";
            OleDbConnection con =
                new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;");
            OleDbDataAdapter da = new OleDbDataAdapter(query, con);
            DataSet aDataObjectSet = new DataSet();
            da.Fill(aDataObjectSet);
            con.Close();
            return aDataObjectSet;
        }

        public void CloseConnection()
        {
            MyConnection.Close();
        }
        public static List<string> RemoveTRXIDFromGCELLFREQ_FREQColumn(string dbFile, string table)
        {
            List<string> xlData = new List<string>();
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;


            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(dbFile, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(table);


            Console.WriteLine(xlWorkSheet.Name);


            xlWorkSheet.Cells[1, 5] = "";
            //xlWorkBook.
            xlWorkBook.SaveAs(dbFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            xlWorkBook.Close(false, false, false);
            xlApp.Quit();

            return xlData;


        }

    }
}