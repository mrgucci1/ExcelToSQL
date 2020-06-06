using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace ExcelToSQL
{
    class getDate
    {
        public string getDateFromFilename(string file)
        {
            //Get last 8 characters from file string
            string temp2 = file;
            temp2 = temp2.Substring(0, temp2.Length - 4);
            temp2 = temp2.GetLast(8);
            //Console.WriteLine(temp);
            //Convert to datetime type
            
            DateTime date = DateTime.ParseExact(temp2, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            string finalDate = date.ToString("MM-dd-yyyy");
            //finalDate = "'" + finalDate + "'";
            return finalDate;
            
            
        }

        
    }
    public static class StringExtension
    {
        public static string GetLast(this string source, int tail_length)
        {
            if (tail_length >= source.Length)
                return source;
            return source.Substring(source.Length - tail_length);
        }
    }
}
