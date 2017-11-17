using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Threading;

namespace ConsoleApplication2

/* What I did in this was create a Console application in VS then use C# code to load data to a database in SQL Server.
 I created database manualy in SQL Server, then table with Columns as in an Excel Sheet. A folder: Log in C:\ drive and error log is saved there.
 Each table is ran separately through nulled and activated TableName and sheetName identifiers*/
{
    class Transfer
    {
        static void Main(string[] args)
        {
            string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
            Console.WriteLine(datetime);
            string LogFolder = @"c:\Log";
            try
            {
                string DatabaseName = "DatasourceSQL";
                string SQLServerName = "Student-PC";
                
                // Table Names to be created and populated in SQL Server Management Studio
                //string TableName = @"Orders";
                string TableName2 = @"Returns";
                //string TableName3= @"Users";
                string SchemaName = @"dbo";
                // Source location of Excel(.xls) file
                string fullFilePath = @"storeSalesExcel.xls";
                //string sheetName = "Orders";
                string sheetName2 = "Returns";
                //string sheetName3 = "Users";

                SqlConnection SQLConnection = new SqlConnection();
                SQLConnection.ConnectionString = "Data Source = "
                    + SQLServerName + "; Initial Catalog = "
                    + DatabaseName + "; "
                    + "Integrated Security = true";

                string ConStr;
                string HDR;
                HDR = "YES";
                ConStr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = "
                    + fullFilePath + "; Extended Properties =\"Excel 12.0; HDR = "
                    + HDR + ";IMEX = 0\"";
                OleDbConnection Conn = new OleDbConnection(ConStr);
                Conn.Open();
               
                //OleDbCommand oconn = new OleDbCommand("select * from [" + sheetName + sheetName2 + sheetName3 + "$]", Conn);
                OleDbCommand oconn = new OleDbCommand("select * from [" + sheetName2 + "$]", Conn);

                OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                DataTable dt = new DataTable();
                adp.Fill(dt);
                Conn.Close();
                
                // Create tables
                SQLConnection.Open();
                using (SqlCommand command = new SqlCommand("CREATE TABLE Returns(Region char(250), Manager char(250));", SQLConnection))
                command.ExecuteNonQuery();

                //using (SqlCommand command = new SqlCommand("CREATE TABLE Orders(Region char(100), Manager char(100), Returns(Region char(100), Manager char(100), Users(Region char(100), Manager char(100));", SQLConnection))
                //command.ExecuteNonQuery(); Previous Code for creation off three tables..

                // establishing table structure in SQL Server Management Studio
                using (SqlBulkCopy BC = new SqlBulkCopy(SQLConnection))
                {

                    BC.DestinationTableName = SchemaName + "." + TableName2;
                    foreach (var column in dt.Columns)
                        BC.WriteToServer(dt);
                    Console.WriteLine("Database {0} table {1} is Loaded with data from .xls file", DatabaseName, BC.DestinationTableName);
                    Thread.Sleep(30000);
                }
                
                SQLConnection.Close();
                

            }
            catch (Exception exception)
            {
                using (StreamWriter sw = File.CreateText(LogFolder
                    + "\\" + "ErrorLog_" + datetime + ".log"))
                {
                    sw.WriteLine(exception.ToString());
                    Console.WriteLine("data to catch");
                }
            }

            
        }
    }
}


                
