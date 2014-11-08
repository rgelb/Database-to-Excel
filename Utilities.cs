using System;
using System.Data;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace DatabaseToExcel
{
    public class Utilities
    {
        public static void RenderDataTableOnXlSheet(DataTable dt, Excel.Worksheet xlWk, string[] columnNames, string[] fieldNames)
        {
            Excel.Range rngExcel = null;
            Excel.Range headerRange = null;

            try
            {
                // render the column names (e.g. headers)
                for (int i = 0; i < columnNames.Length; i++)
                    xlWk.Cells[1, i + 1] = columnNames[i];


                if (dt.Rows.Count > 0)
                {

                    // for each column, create an array and set the array 
                    // to the excel range for that column.
                    for (int i = 0; i < fieldNames.Length; i++)
                    {
                        var clnDataString = new string[dt.Rows.Count,1];
                        var clnDataInt = new int[dt.Rows.Count,1];
                        var clnDataDouble = new double[dt.Rows.Count,1];

                        //string columnLetter = char.ConvertFromUtf32("A".ToCharArray()[0] + i);
                        string columnLetter = IndexToExcelColumnName(i);

                        rngExcel = xlWk.get_Range(columnLetter + "2", Missing.Value);
                        rngExcel = rngExcel.get_Resize(dt.Rows.Count, 1);

                        string dataTypeName = dt.Columns[fieldNames[i]].DataType.Name;

                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (fieldNames[i].Length > 0)
                            {
                                if (!dt.Rows[j].IsNull(i))
                                {
                                    switch (dataTypeName)
                                    {
                                        case "Int32":
                                            clnDataInt[j, 0] = Convert.ToInt32(dt.Rows[j][fieldNames[i]]);
                                            break;
                                        case "Double":
                                        case "Decimal":
                                            clnDataDouble[j, 0] = Convert.ToDouble(dt.Rows[j][fieldNames[i]]);
                                            break;
                                        case "DateTime":
                                            if (fieldNames[i].ToLower().Contains("time"))
                                                clnDataString[j, 0] = Convert.ToDateTime(dt.Rows[j][fieldNames[i]]).ToShortTimeString();
                                            else
                                                clnDataString[j, 0] = Convert.ToDateTime(dt.Rows[j][fieldNames[i]]).ToString();

                                            break;
                                        default:
                                            clnDataString[j, 0] = dt.Rows[j][fieldNames[i]].ToString();
                                            break;
                                    }
                                }
                                else
                                    clnDataString[j, 0] = "NULL";
                            }
                            else
                                clnDataString[j, 0] = string.Empty;
                        }

                        // set values in the sheet wholesale.
                        if (dataTypeName == "Int32")
                            rngExcel.set_Value(Missing.Value, clnDataInt);
                        else if (dataTypeName == "Double")
                            rngExcel.set_Value(Missing.Value, clnDataDouble);
                        else
                            rngExcel.set_Value(Missing.Value, clnDataString);
                    }
                }


                // figure out the letter of the last column (supports 1 letter column names)
                //string lastColumn = char.ConvertFromUtf32("A".ToCharArray()[0] + columnNames.Length - 1);
                string lastColumn = IndexToExcelColumnName(columnNames.Length - 1);

                // make the header range bold
                headerRange = xlWk.get_Range("A1", lastColumn + "1");
                headerRange.Font.Bold = true;

                // autofit for better view
                xlWk.Columns.AutoFit();

            }
            finally
            {
                ReleaseComObject(headerRange);
                ReleaseComObject(rngExcel);
            }
        }

        public static void ReleaseComObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }


        /// <summary>
        /// Converts an index to an Excel column name
        /// </summary>
        /// <param name="index">Assumes 0 based index</param>
        /// <returns></returns>
        private static string IndexToExcelColumnName(int index)
        {
            index++;
            int dividend = index;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
        
    }
}
