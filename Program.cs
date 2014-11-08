using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace DatabaseToExcel
{
    class Program
    {
        private static readonly AppArgs parsedArgs = new AppArgs();

        static void Main(string[] args)
        {
            if (Parser.ParseArgumentsWithUsage(args, parsedArgs))
            {
                try
                {
                    Log("Building Connection String");
                    string connString = GetConnectionString(parsedArgs);
                    Log("Fetching Data");
                    DataSet ds = GetData(parsedArgs.queryFile, connString);
                    Log("Creating Excel File");
                    CreateExcelFile(parsedArgs, ds);
                }
                catch (Exception ex)
                {
                    Log("Error occured: " + ex.Message);
                    Environment.Exit(1);
                }
            }
            else
            {
                Log("Invalid arguments");
                Environment.Exit(1);
            }
        }

        private static void Log(string msg)
        {
            Console.WriteLine(msg);
            Trace.WriteLine(msg);
            Debug.WriteLine(msg);
        }

        private static void CreateExcelFile(AppArgs appArgs, DataSet ds)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null; 

            var sheetNames = new List<string>();
            if (appArgs.sheetFile.Length > 0 && File.Exists(appArgs.sheetFile))
                sheetNames = File.ReadAllLines(appArgs.sheetFile).ToList();

            try
            {
                app = new Excel.Application {Visible = false};
                workbook = app.Workbooks.Add(1);
                 //   = (Excel.Worksheet)workbook.Sheets[1];

                for (int i = 0; i < ds.Tables.Count - 1; i++)
                    workbook.Sheets.Add();     

                // now name them
                for (int i = 0; i < sheetNames.Count; i++)
                {
                    if (workbook.Sheets.Count < i) break;
                    worksheet = (Excel.Worksheet)workbook.Sheets[i + 1];
                    worksheet.Name = sheetNames[i];
                }

                // now populate the spreadsheet
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    //var columnNames = new List<string>();
                    DataTable dt = ds.Tables[i];
                    worksheet = (Excel.Worksheet)workbook.Sheets[i + 1];

                    List<string> columnNames = dt.Columns.Cast<DataColumn>().Select(cln => cln.ColumnName).ToList();

                    //foreach (DataColumn item in dt.Columns)
                    //    columnNames.Add(item.ColumnName);

                    Utilities.RenderDataTableOnXlSheet(dt, worksheet, columnNames.ToArray(), columnNames.ToArray());
                }

                // delete output file if exists
                if (File.Exists(appArgs.outputFile)) File.Delete(appArgs.outputFile);

                workbook.SaveAs(appArgs.outputFile);
                
                app.Quit();
                
            }
            catch (Exception e)
            {
                Console.Write("Error: " + e.Message);
            }
            finally
            {
                Utilities.ReleaseComObject(worksheet);
                Utilities.ReleaseComObject(workbook);
                Utilities.ReleaseComObject(app);
            }

        }

        private static string GetConnectionString(AppArgs args)
        {
            return string.Format("data source={0};initial catalog={1};user id={2};password={3}",
                    args.server, args.database, args.user, args.password);
        }

        private static DataSet GetData(string queryFile, string connString)
        {
            string queryContents = File.ReadAllText(queryFile);
            return ExecuteDataSet(connString, queryContents);
        }

        protected static DataSet ExecuteDataSet(string connectionString, string commandText)
        {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = commandText;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandTimeout = 120;

                    using (var da = new SqlDataAdapter(cmd))
                    {
                        var ds = new DataSet();
                        da.Fill(ds);

                        return ds;
                    }
                }
            }
        }
    }
}
