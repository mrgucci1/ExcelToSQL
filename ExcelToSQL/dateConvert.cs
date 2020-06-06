using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSQL
{
    class dateConvert
    {
        public string dateConvertStart(string convert, string fieldName)
        {
            try
            {
                //convert to double
                convert = convert.Replace("'", "");
                convert = convert.Trim();
                double temp = double.Parse(convert);
                //convert to datetime
                DateTime conv = DateTime.FromOADate(temp);
                //convert back to string in our format
                convert = conv.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch
            {
                if(convert != "NULL")
                    Console.WriteLine($"Failed to Convert a date: {fieldName}");
                Console.WriteLine($"Date was passed in as: {convert}");
            }
            if (convert == "NULL")
                return convert;
            else
            {
                convert = "'" + convert + "'";
                return convert;
            }
            
        }
        public string dateTimeConvertStart(string convert, string fieldName)
        {
            try
            {
                //convert to long
                convert = convert.Replace("'", "");
                convert = convert.Trim();
                double temp = double.Parse(convert);
                //convert to datetime
                DateTime conv = DateTime.FromOADate(temp);
                //convert back to string in our format
                convert = conv.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            }
            catch
            {
                if (convert != "NULL")
                    Console.WriteLine($"Failed to Convert a date: {fieldName}");
                Console.WriteLine($"Date was passed in as: {convert}");
            }
            if (convert == "NULL")
                return convert;
            else
            {
                convert = "'" + convert + "'";
                return convert;
            }
        }
    }
}
