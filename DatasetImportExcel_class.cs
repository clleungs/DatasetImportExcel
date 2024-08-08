using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.Remoting;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace DatasetImportExcel
{

    public static class Globals
    {
        public static bool verbose = false;
        public static string Customer = "";
        public static string Quarter = "";
        public static int HeaderRow = 0;
        public static string po = "";
        public static string po_nbr = "";
        public static string po_date_str = ""; 
        public static int QtyColCnt = 0;
        public static int Product_descCol = 0; 
        public static int UmCol = 0;
        public static string [] QuarterLabel = new string[3] { "", "", "" };
        public static int[] QtyCol = new int[3]; //0,1,2
        

        const string GREEN = "\x1B[92m";
        const string CYAN = "\x1B[96m";
        const string YELLOW = "\x1B[93m";
        const string YELLOW2 = "\x1B[1;33m";   // an ANSI escape code with Yellow Bold.
        const string ORANGE = "\x1B[38;2;255;165;0m";
        const string RED = "\x1B[91m";
        const string TIFFANY_BLUE = "\u001b[38;5;123m";
        const string MAGENTA = "\x1B[95m";
        const string BOLD = "\u001b[1m";
        const string ITALIC = "\u001b[3m";
        const string STRIKETHROUGH = "\u001b[9m";
        const string GRAY = "\x1B[90m";
        const string RESET = "\x1B[0m\x1B[037m";

        // statiic Constructor with default values
        static Globals()
        {

        }
        public static void ListQtyCol()
        {
            Console.WriteLine(YELLOW + "Listing QtyCol Column Position (0-column A, 1-column B,....)" + RESET); 
            for (int i = 0; i < 3; i++)
            {
                Console.WriteLine(Globals.QtyCol[i]);   
            }
        }
        public static void ListCustomer()
        {
            var className = nameof(Globals);
            var st = new StackTrace();
            var sf = st.GetFrame(0);
            if (verbose )
            {
                Console.WriteLine("{1}Procedure : {3} {0}{2}", sf?.GetMethod(), GRAY, RESET, className);
            }
            Console.WriteLine(YELLOW + "Listing Customer" +  RESET);
            Console.WriteLine("Customer : {0}", Customer);
            Console.WriteLine("Quarter  : {0}", Quarter);
        }
    }
    //////////////////////////////////////////////////////////////////////////////////////////////////
    partial class DatasetImportExcel
    {
      //NULL

    }
}
