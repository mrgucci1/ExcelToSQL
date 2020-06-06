using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text.RegularExpressions;

namespace ExcelToSQL
{
    class excelMaster
    {
        public object[,] excelMasterStart(string path)
        {
            //Create Excel Objects 
            Excel.Application ExcelApp;
            Excel._Workbook ExcelWorkbook;
            Excel._Worksheet ExcelSheet;
            //Initilize ExcelApp
            ExcelApp = new Excel.Application();
            //Open current Excel file
            Console.WriteLine($"Opening: {path}");
            ExcelWorkbook = ExcelApp.Workbooks.Open(path);
            //Set Active Sheet
            ExcelSheet = ExcelWorkbook.Sheets[1];
            //ExcelApp.Visible = true;
            Excel.Range ExcelRange = ExcelSheet.UsedRange;
            int rowCount = ExcelRange.Rows.Count;
            int colCount = ExcelRange.Columns.Count;
            rowCountF = rowCount;
            colCountF = colCount;
            object[,] excelValues = (object[,])ExcelRange.Value2;
            //Close and release all COM objects, quit excel
            Marshal.ReleaseComObject(ExcelRange);
            Marshal.ReleaseComObject(ExcelSheet);
            //close and release
            ExcelWorkbook.Close();
            Marshal.ReleaseComObject(ExcelWorkbook);
            //quit and release
            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelApp);
            return excelValues;
        }
        public int rowCountF { get; set; }
        public int colCountF { get; set; }
        public object[,] filterExcel(object[,] excelValues, int[] numArray, int startingRow)
        {
            for (int rows = startingRow; rows < rowCountF + 1; rows++)
            {
                //Cleanup excel data, remove null cells and replace single quotes
                for (int cols = 1; cols < colCountF + 1; cols++)
                {
                    if (excelValues[rows, cols] == null || excelValues[rows,cols].ToString() == "")
                    {
                        excelValues[rows, cols] = "NULL";
                    }
                    if (excelValues[rows, cols].ToString().Contains("NULL"))
                    {

                    }
                    else
                    {
                        excelValues[rows, cols] = excelValues[rows, cols].ToString().Replace("'", "''");
                        excelValues[rows, cols] = excelValues[rows, cols].ToString().Trim();
                        //Console.WriteLine($"Replaced Quotes on Row: {rows} and Cols: {cols}");
                    }

                }
                for (int cols = 1; cols < colCountF + 1; cols++)
                {
                    //skip rows that do not need have single quotes
                    if (numArray.Contains(cols))
                        continue;
                    if (excelValues[rows, cols].ToString().Contains("NULL"))
                    {

                    }
                    else
                    {
                        excelValues[rows, cols] = excelValues[rows, cols].ToString().Trim();
                        excelValues[rows, cols] = "'" + excelValues[rows, cols] + "'";

                        //Console.WriteLine($"Added Quotes Around on Row: {rows} and Cols: {cols}");
                    }
                }
            }
            return excelValues;
        }
    }
}
