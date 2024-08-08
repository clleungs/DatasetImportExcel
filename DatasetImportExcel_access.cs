using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace DatasetImportExcel
{
    partial class DatasetImportExcel
    {
        private static OleDbConnection GetConnection(string filename, bool openIt)
        {
            // but always ignores the first row of Excel
            var st = new StackTrace();
            var sf = st.GetFrame(0);
            if (Globals.verbose)
            {
                Console.WriteLine("{1}Procedure : {0}{2}", sf?.GetMethod() + " FileName:" + filename, GRAY, RESET);
            }
            // If your data has no header row, change HDR=NO
            var c = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{filename}';Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\" ");
            if (openIt)
                c.Open();
            return c;
        }

        private static OleDbConnection GetConnection2(string filename, bool openIt)
        {
            // No header
            var st = new StackTrace();
            var sf = st.GetFrame(0);
            if (Globals.verbose)
            {
                Console.WriteLine("{1}Procedure : {0}{2}", sf?.GetMethod() + " FileName:" + filename, GRAY, RESET);
            }
            // If your data has no header row, change HDR=NO
            var c = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{filename}';Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\" ");
            if (openIt)
                c.Open();
            return c;
        }

        //////////////////////////////////////////////////////////////////////////////////
        private static DataSet GetExcelFileAsDataSet(OleDbConnection conn)
        {
            var st = new StackTrace();
            var sf = st.GetFrame(0);
            if (Globals.verbose)
            {
                Console.WriteLine("{1}Procedure : {0}{2}", sf?.GetMethod() + " OleDBConn:" + conn, GRAY, RESET);
            }
            var sheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new[] { default, default, default, "TABLE" });
            var ds = new DataSet();
            foreach (DataRow r in sheets.Rows)
            {
                string tableName = r["TABLE_NAME"].ToString();
                if (!tableName.EndsWith("_xlnm#_FilterDatabase"))
                {
                    ds.Tables.Add(GetExcelSheetAsDataTable(conn, r["TABLE_NAME"].ToString()));
                }
            }
            return ds;
        }
        //////////////////////////////////////////////////////////////////////////////////
        private static DataTable GetExcelSheetAsDataTable(OleDbConnection conn, string sheetName)
        {

            var st = new StackTrace();
            var sf = st.GetFrame(0);
            if (Globals.verbose)
            {
                Console.WriteLine("{1}Procedure : {0}{2}", sf?.GetMethod() + " SheetName:" + sheetName, GRAY, RESET);
            }

            using (var da = new OleDbDataAdapter($"select * from [{sheetName}]", conn))
            {
                var dt = new DataTable() { TableName = sheetName.TrimEnd('$') };
                da.Fill(dt);
                if (Globals.verbose)
                {
                    Console.WriteLine("\t{1}Sheetname {0}{2}", sheetName, GRAY, RESET);
                    Console.WriteLine("\t{1}Row Count:{0}{2}", dt.Rows.Count, GRAY, RESET);
                }
                return dt;
            }
        }
        //////////////////////////////////////////////////////////////////////////////////
        static bool CheckMissingFieldsInFirstRow(String fn, DataSet dataSet, string sheetName, string[] fields)
        {
            // SheetName should not have -
            bool found = false;

            DataTable targetSheet = dataSet.Tables[sheetName];
            if (Globals.verbose)
            {
                Console.WriteLine("sheetName= {0}", sheetName);
            }
            if (targetSheet != null && targetSheet.Rows.Count > 0)
            {
                DataRow firstRow = targetSheet.Rows[0];

                var missingFields = fields.Where(field => !targetSheet.Columns.Contains(field) || firstRow.IsNull(field)).ToList();

                if (missingFields.Any())
                {
                    Console.WriteLine("Missing fields {0} of sheet[{1}] ", fn, sheetName);
                    foreach (var missingField in missingFields)
                    {
                        Console.WriteLine(missingField);
                    }
                }
                else
                {
                    Console.WriteLine("All specified fields found {0} sheet[{1}]", fn, sheetName);
                    found = true;
                }
            }
            else
            {
                Console.WriteLine(RED + "ERROR: Sheet with name '" + sheetName + "' not found or empty in the DataSet." + RESET);
            }
            return found;
        }
        //////////////////////////////////////////////////////////////////////////////////////////////////

        static void LoadExcelWithHDR(string fn)
        {
            if (File.Exists(fn))
            {
                using (var c2 = GetConnection(fn, true))
                    dsExcel = GetExcelFileAsDataSet(c2);
                dsExcel.EnforceConstraints = false;

            } 
        }

        static void LoadExcelNoHDR(string fn)
        {
            if (File.Exists(fn))
            {
                using (var c2 = GetConnection2(fn, true))
                    dsExcel = GetExcelFileAsDataSet(c2);
                dsExcel.EnforceConstraints = false;

            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////
        // Function to convert an integer to a column letter (A, B, C, ...)
        public static string ConvertToLetter(int columnNumber)
        {
            int dividend = columnNumber + 1;
            string columnName = "";

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        //////////////////////////////////////////////////////////////////////////////////////////////////
        static void RenameColumnWithCheck(DataTable table, int columnIndex, string newColumnName, Type dataType)
        {
            int index = table.Columns.IndexOf(newColumnName);
            if (index != -1)
            {
                int suffix = 1;
                string uniqueName = newColumnName + "-" + suffix;

                while (table.Columns.Contains(uniqueName))
                {
                    suffix++;
                    uniqueName = newColumnName + "-" + suffix;
                }

                newColumnName = uniqueName;
            }

            if (columnIndex < table.Columns.Count)
            {
                table.Columns[columnIndex].ColumnName = newColumnName;
                if (dataType != null)
                {
                    table.Columns[columnIndex].DataType = dataType;
                }
            }
        }
    }
}
