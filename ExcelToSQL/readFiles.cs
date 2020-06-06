using System;
using System.Collections.Generic;
using System.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ExcelToSQL
{
    public class readFiles
    {
        //Used parts of: https://docs.microsoft.com/en-us/dotnet/api/system.io.directory.getfiles?view=netframework-4.8
        public string[] Main()
        {
            string path = Directory.GetCurrentDirectory();
            string[] empty = new string[1];
            if (File.Exists(path))
            {
                // This path is a file
                Console.WriteLine("Filepath was to a file");
                return empty;
            }
            else if (Directory.Exists(path))
            {
                // This path is a directory
                string[] fileEntries = ProcessDirectory(path);
                return fileEntries;
            }
            else
            {
                Console.WriteLine("{0} is not a valid file or directory.", path);
                return empty;
            }
        }
        // Process all files in the directory passed in
        public string[] ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            return fileEntries;
        }
    }
}
