using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

/* 2024/1111dddfd */

namespace DatasetImportExcel
{
    public static class StringExtensions
    {
        public static string Repeat(string value, int count)
        {
            return new StringBuilder(value.Length * count).Insert(0, value, count).ToString();
        }
        public static string Repeat(this char c, int n)
        {
            return new String(c, n);
        }
        public static string xfill(this string _this, string b, int count)
        {
            string cstr = "";
            if (_this.Length + 1 >= count)
            {
                cstr = _this;
            }
            else
            {
                cstr = _this + " " + ".".PadRight(count - _this.Length - 1, '.'); ;
            }
            return cstr;
        }
    }
    //////////////////////////////////////////////////////////////////////////////////////////////////

    partial class DatasetImportExcel
    {
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

        static void Help()
        {
            String msg;
            // System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            string prog = System.Reflection.Assembly.GetExecutingAssembly().ToString();
            string[] vprog = prog.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            // msg = Properties.readme.Default.help.Replace("|", "\n");

            var asm = Assembly.GetExecutingAssembly();   // Using  System.Reflection;
            var basm = asm.CustomAttributes.FirstOrDefault(a => a.AttributeType == typeof(TargetFrameworkAttribute));
            var strFramework = basm.NamedArguments[0].TypedValue.Value;

            msg = "";
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\nHelp ℹ️ Load Excel using Access Engine Check Excel File Order for which Customer");
            Console.ResetColor();
            String vfmt01 = "{0,-70} {1}";
            msg = "...." + "\n";
            Console.WriteLine(msg);

            msg = string.Format(vfmt01, vprog[0], "")
              + "\n" + string.Format(vfmt01, vprog[0] + " -f \"Advance One - Order Template Q3 2024.xlsx\" -xlsx", "#")
              + "\n" + string.Format(vfmt01, vprog[0] + " -c qube -f \"202312 - Qube Net Romania MAR 2024 Order template.xlsx\" -xlsx", "#")
              + "\n" + GREEN + "Source:" + RESET
              + "\n" + @"F:\prog-c\Utilities\Dataset\DatasetImportExcel\bin\Debug" + "\tImplementation: " + CYAN + strFramework + " x64" + RESET
            ;
            Console.WriteLine(msg);
            System.Environment.Exit(-1);
        }
       
    }
}
