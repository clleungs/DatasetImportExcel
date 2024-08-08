using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Net.NetworkInformation;
/* ddd*/
namespace DatasetImportExcel
{
    public static class Globals
    {
        public static string fn = "AdvanceOne-OrderTemplateQ32024.xlsx"; 
        public static bool verbose = false;
    }

    partial class DatasetImportExcel
    {
        public static DataSet dsExcel;
        static void Main(string[] args)
        {
            string fn = "AdvanceOne-OrderTemplateQ32024.xlsx";
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-h")
                {
                    Help();
                }
                else if (args[i] == "-f")
                {
                    fn = args.Length > i + 1 ? args[i + 1] : "";
                }
            }

            LoadExcel(fn);
            Console.WriteLine("Loaded .. {0}", fn);

            // Rename the columns
            //dsExcel.Tables[0].Columns["F1"].ColumnName = "Product series";
            DataTable table2 = dsExcel.Tables[0];

            // Renaming columns and checking for existing names
            RenameColumnWithCheck(table2,0, "Product series" ,null);
            RenameColumnWithCheck(table2,1, "Item", null);
            RenameColumnWithCheck(table2,2, "Product Description", null);
            RenameColumnWithCheck(table2,3, "Product Description", null);
            RenameColumnWithCheck(table2,4, "QTY/CTN", null);
            RenameColumnWithCheck(table2, 5, "UM", null);
            RenameColumnWithCheck(table2, 6, "Stock on Hand", null);
            RenameColumnWithCheck(table2, 7, "Coming", null);
            RenameColumnWithCheck(table2, 8, "Allocated", null);
            RenameColumnWithCheck(table2, 9, "3 month sales", null);

        
            // Assuming dsExcel is your DataSet
            foreach (DataTable table in dsExcel.Tables)
            {
                Console.WriteLine($"Column Headers for Table: {table.TableName}");
                int colIndex = 0;
                foreach (DataColumn column in table.Columns)
                {
                    string columnName = ConvertToLetter(colIndex);
                    Console.WriteLine($"{columnName} {colIndex} ({column.DataType}) {column.ColumnName} ");
                    colIndex++;
                }
                Console.WriteLine(); // Add a blank line between tables
            }


            // Rename field 
            // Assuming dsExcel is your DataSet
            int rowIndex = 9; // Row 10 (0-based index)
            int startColumnIndex = 0; // Starting column index
            int endColumnIndex = 20; // Ending column index (inclusive)

            if (dsExcel.Tables.Count > 0 && dsExcel.Tables[0].Rows.Count > rowIndex)
            {
                Console.WriteLine($"Values for Row 10 (0-based index) from Column 0 to 20:");

                // Iterate through the columns within the specified range
                for (int columnIndex = startColumnIndex; columnIndex <= endColumnIndex && columnIndex < dsExcel.Tables[0].Columns.Count; columnIndex++)
                {
                    string columnName = ConvertToLetter(columnIndex);
                    Console.WriteLine($"{columnName} Column {columnIndex}: {dsExcel.Tables[0].Rows[rowIndex][columnIndex]}");
                }
            }
            else
            {
                Console.WriteLine("Row 10 or the specified columns are out of bounds.");
            }
        }

    }
}
