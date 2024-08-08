using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Web;


/* 2024/08/05 review https://stackoverflow.com/questions/45350278/c-sharp-reading-excel-file-ignores-1st-row
                     Change "HDR" to "No"
 * 2024/07/31 Review 
 * 2024/07/28 
 */

namespace DatasetImportExcel
{
    partial class DatasetImportExcel
    {
        public static DataSet dsExcel;
        static void HandleAAA(string fn)
        {
            // Microsoft Interop Excel skip row with RowMerged 
            int MaxRow;

            LoadExcelWithHDR(fn);

            // Assigning a variable to a reference in C# does not double the memory consumption.
            // This means that both excelPO and dsExcel.Tables[0] will point to the same object in memory.
            var excelPO = dsExcel.Tables[0];

            Console.WriteLine("Loaded ........ : {1}{0}{2}", fn, YELLOW, RESET);
            MaxRow = excelPO.Rows.Count;

            Console.WriteLine("Sheet Name  ... : {1}{0}{2}", excelPO.TableName, YELLOW, RESET);
            Console.WriteLine("Column Count .. : {1}{0}{2}", excelPO.Columns.Count, YELLOW, RESET);
            Console.WriteLine("{0}Checking File Layout:{1}", YELLOW, RESET);
            int endColumnIndex = excelPO.Columns.Count; // Ending column index (inclusive)
            // In C#, arrays, lists, and other collections are zero-indexed
            for (int jc = 0; jc < MaxRow; jc++)
            {
                if (excelPO.Rows[jc][0] != DBNull.Value && excelPO.Rows[jc][0].ToString().StartsWith("Order Template"))
                {
                    Console.WriteLine("{3}cells({0},{1})={2}{4}", jc, 0, excelPO.Rows[jc][0], GRAY, RESET);
                }

                if (excelPO.Rows[jc][2] != DBNull.Value && excelPO.Rows[jc][2].ToString().StartsWith("Order Cycle"))
                {
                    Console.WriteLine("{3}cells({0},{1})={2}{4}", jc, 2, excelPO.Rows[jc][2], GRAY, RESET);
                    Console.WriteLine("{3}cells({0},{1})={2}{4}", jc, 1, excelPO.Rows[jc][1], GRAY, RESET);
                    Globals.Customer = excelPO.Rows[jc][2].ToString();
                    Globals.Quarter = excelPO.Rows[jc][1].ToString();
                }
                if (excelPO.Rows[jc][1] != DBNull.Value && excelPO.Rows[jc][1].ToString().StartsWith("Item") &&
                   excelPO.Rows[jc][2] != DBNull.Value && excelPO.Rows[jc][2].ToString().StartsWith("Product Description"))
                {
                    Console.WriteLine("{3}cells({0},{1})={2}{4}", jc, 1, excelPO.Rows[jc][1], GRAY, RESET);
                    Console.WriteLine("{3}cells({0},{1})={2}{4}", jc, 2, excelPO.Rows[jc][2], GRAY, RESET);
                    Globals.HeaderRow = jc;
                    Globals.QtyColCnt = 0;
                    Console.WriteLine("{0}Header Rows Row,Column start from 0{1}", YELLOW, RESET);
                    Console.WriteLine("{0,8} {1,8} {2}", "Row", "Column", "Column label");
                    Console.WriteLine("{0,8} {1,8} {2}", '-'.Repeat(8), '-'.Repeat(8), '-'.Repeat(35));
                    for (int jd = 0; jd < excelPO.Columns.Count; jd++)
                    {
                        // Check header : Order Qty
                        Console.WriteLine("{0,8} {1,8} {2}", jc, jd, excelPO.Rows[jc][jd].ToString());
                        if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString() == "Order Qty")
                        {
                            Globals.QtyCol[Globals.QtyColCnt] = jd;
                            Globals.QuarterLabel[Globals.QtyColCnt] = excelPO.Rows[jc - 1][jd].ToString();
                            if (Globals.QtyColCnt < 3) Globals.QtyColCnt++;
                        }
                        if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString() == "Product Description")
                        {
                            Globals.Product_descCol = jd;
                        }
                        if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString() == "UM")
                        {
                            Globals.UmCol = jd;
                        }
                    }
                }

            }
            Globals.ListQtyCol();

            /*
              if (dsExcel.Tables.Count > 0)
              {
                  DataTable table = dsExcel.Tables[0];

                  // Assuming column index 0 for column A
                  int rowNum = 1; // Initialize the row number counter
                  foreach (DataRow row in table.Rows)
                  {
                      object columnAValue = row[0]; // Index 0 for column A
                      object columnBValue = row[1]; // Index 1 for column B
                      object columnCValue = row[2]; // Index 2 for column C
                      string Quarter = "";
                      if (columnAValue != DBNull.Value)
                      {
                          string columnAStringValue = columnAValue.ToString();

                          if (Regex.IsMatch(columnAStringValue, "Order Template", RegexOptions.IgnoreCase))
                          {
                              Console.WriteLine("Customer cells({3},0): {1}{0}{2}", columnAStringValue, YELLOW, RESET, rowNum);
                              Regex regex = new Regex("Order Template - (.*)", RegexOptions.IgnoreCase);
                              Match match = regex.Match(columnAStringValue);
                              if (match.Success)
                              {
                                  Globals.Customer = match.Groups[1].Value.Trim();
                              }
                          }
                          // Additional processing for columnAStringValue
                      }
                      if (columnCValue != DBNull.Value)
                      {
                          string columnCStringValue = columnCValue.ToString();
                          if (columnCStringValue.Contains("Order Cycle"))
                          {
                              if (columnBValue != DBNull.Value)
                              {
                                  Quarter = columnBValue.ToString();
                              }
                              Console.WriteLine("Customer cells({3},2): {1}{0} -> {2}Quarter{1} {4}{2}", columnCStringValue, YELLOW, RESET, rowNum, Quarter);
                              Globals.Quarter = Quarter;
                          }
                      }
                      rowNum++; // Increment the row number after processing each row
                  } 
              } */

        }
        //////////////////////////////////////////////////////////////////////////////////////////////////
        static void WorkSheetPrint()
        {
            // When you assign an object to another variable in C#, you are not creating a copy of the object.
            // Instead, you are creating another reference to the same object in memory.
            DataTable table = dsExcel.Tables[0];
            int rowNum = 0; // Initialize the row number counter
            Console.WriteLine();
            foreach (DataRow row in table.Rows)
            {
                if (rowNum == Globals.HeaderRow)
                {
                    // Print Header
                    Console.WriteLine("{0,-18} {1,-50} {2,5} {3,-12} {4,-12} {5,-12} {6,10}", "", "", "", Globals.QuarterLabel[0], Globals.QuarterLabel[1], Globals.QuarterLabel[2], "");
                    Console.WriteLine("{0,-18} {1,-50} {2,5} {3,-12} {4,-12} {5,-12} {6,10}", row[1], "Product Description", "UM", row[Globals.QtyCol[0]], row[Globals.QtyCol[1]], row[Globals.QtyCol[2]], "Sum");
                    Console.WriteLine("{0,-18} {1,-50} {2,5} {3,-12} {4,-12} {5,-12} {6,10}", '-'.Repeat(18), '-'.Repeat(50), '-'.Repeat(5), '-'.Repeat(12), '-'.Repeat(12), '-'.Repeat(12), '-'.Repeat(10));
                }
                // 6 = Stock on Hand , 7 = Coming , 11-13 Order Qty 
                double sum = 0;
                for (int ix = Globals.QtyCol[0]; ix <= Globals.QtyCol[2]; ix++)
                {
                    if (row[ix] != DBNull.Value && row[ix] != null)
                    {
                        if (double.TryParse(row[ix].ToString(), out double value))
                        {
                            sum += value;
                        }
                    }
                }
                if (sum > 0)
                {
                    Console.WriteLine("{0,-18} {1,-50} {2,5} {3,12} {4,12} {5,12} {6,10}",
                         row[1], row[Globals.Product_descCol], row[Globals.UmCol], row[Globals.QtyCol[0]], row[Globals.QtyCol[1]], row[Globals.QtyCol[2]], sum);
                }
                rowNum++;
            }
        }
        //////////////////////////////////////////////////////////////////////////////////////////////////
        static void Main(string[] args)
        {
            string fn = "Advance One - Order Template Q3 2024.xlsx";
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-h")
                {
                    Help();
                }
                else if (args[i] == "-v")
                {
                    Globals.verbose = true;
                }
                else if (args[i] == "-c")
                {
                    Globals.Customer = args.Length > i + 1 ? args[i + 1] : "";
                }
                else if (args[i] == "-f")
                {
                    fn = args.Length > i + 1 ? args[i + 1] : "";
                }
            }
            if (!File.Exists(fn))
            {
                Console.WriteLine("ERROR: File not exist {0}", fn);
                System.Environment.Exit(-1);
            }
            // check file Name 
            if (Globals.Customer == "")
            {
                string fileName = Path.GetFileName(fn);
                if (fileName.Contains("Qube Net"))
                {
                    Globals.Customer = "qube";
                }
                else if (fileName.Contains("Advance Audio") || fileName.Contains("Advance One") || fileName.Contains("CDL"))
                {
                    Globals.Customer = "AAA";
                }
            }

            if (Globals.Customer == "qube")
            {
                Console.WriteLine("{0}First Header also have data{1}", YELLOW, RESET);
                HandleQube(fn);
            }

            else if (Globals.Customer == "AAA")
            {
                HandleAAA(fn);
                Globals.ListCustomer();
                WorkSheetPrint();

            }
        }   // Main

        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        static void HandleQube(String fn)
        {
            int MaxRow;
            LoadExcelNoHDR(fn); // Load excel file 
            var excelPO = dsExcel.Tables["'Order Fulfilment$'"];
            Console.WriteLine("Loaded ........ : {1}{0}{2}", fn, YELLOW, RESET);
            Console.WriteLine("Table Name .... : {1}{0}{2}", excelPO.TableName, YELLOW, RESET);
            MaxRow = excelPO.Rows.Count;
            string altEnter = $"{(char)3}\n"; // Alt + Enter sequence
            for (int jc = 0; jc < 10; jc++)
            {
                for (int jd = 0; jd <= 13; jd++)
                {
                    int krow = jc + 1;
                    string col = ConvertToLetter(jd);
                    Console.Write("{5}({0},{1}={2}{3}){6}{7}{4}{6}", jc, jd, col, krow, (excelPO.Rows[jc][jd]).ToString().Replace(altEnter, GREEN + " * " + RESET), GRAY, RESET, ORANGE);
                    if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString() == "PO No.:")
                    {
                        Globals.po_nbr = excelPO.Rows[jc ][jd + 1].ToString();
                    }
                    if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString()  == "PO Date:") {
                        Globals.po_date_str = excelPO.Rows[jc][jd + 1].ToString();
                        Globals.Customer= excelPO.Rows[jc][jd + 3].ToString();
                    }
                    if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString() == "NEW ORDER REQUEST")
                    {
                        Globals.QtyCol[0] = jd;
                        Globals.QuarterLabel[0] = "NEW ORDER REQUEST";
                    }
                }
                Console.WriteLine();
            }
            Console.WriteLine("NEW ORDER REQUEST Column Index= {0}", Globals.QtyCol[Globals.QtyColCnt]);

            for (int jc = 0; jc < MaxRow; jc++)
            {
                if (excelPO.Rows[jc][1] != DBNull.Value && excelPO.Rows[jc][1].ToString().StartsWith("Product name") &&
                excelPO.Rows[jc][2] != DBNull.Value && excelPO.Rows[jc][2].ToString().StartsWith("Product series"))
                {
                    Globals.HeaderRow = jc;
                    for (int jd = 0; jd < excelPO.Columns.Count; jd++)
                    {
                        // Check header : Order Qty
                        Console.WriteLine("{0,8} {1,8} {2}", jc, jd, excelPO.Rows[jc][jd].ToString());

                        if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString() == "Product Description")
                        {
                            Globals.Product_descCol = jd;
                        }
                        if (excelPO.Rows[jc][jd] != DBNull.Value && excelPO.Rows[jc][jd].ToString() == "UM")
                        {
                            Globals.UmCol = jd;
                        }
                    }
                }
            }
            // print Worksheet here 
            DataTable table = dsExcel.Tables["'Order Fulfilment$'"];
            int rowNum = 0;
            Console.WriteLine("PO        : {0}",Globals.po_nbr);
            Console.WriteLine("PO date   : {0}",Globals.po_date_str);
            Console.WriteLine("Custromer : {0}",Globals.Customer);
            foreach (DataRow row in table.Rows)
            {
                if (rowNum == Globals.HeaderRow)
                {
                    Console.WriteLine("Header {0}", Globals.HeaderRow);
                    Console.WriteLine("{0,-18} {1,6} {2,15}","", "", Globals.QuarterLabel[0]);
                    Console.WriteLine("{0,-18} {1,6} {2,15}", row[3], "UM","Qty"); 
                    Console.WriteLine("{0,-18} {1,6} {2,15}", '-'.Repeat(18), '-'.Repeat(6), '-'.Repeat(15));
                }
                double sum = 0;
                for (int ix = Globals.QtyCol[0]; ix <= Globals.QtyCol[0]; ix++)
                {
                    if (row[ix] != DBNull.Value && row[ix] != null)
                    {
                        if (double.TryParse(row[ix].ToString(), out double value))
                        {
                            sum += value;
                        }
                    }
                }
                if (sum > 0)
                {
                    Console.WriteLine("{0,-18} {1,6} {2,15}", row[3], row[6],  sum);
                }
                rowNum++;
            }
        }
    }
}