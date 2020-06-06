using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text.RegularExpressions;

namespace ExcelToSQL
{
    class excelGenerate
    {
        //Reference to classes
        readFiles rf = new readFiles();
        filterClass fc = new filterClass();
        excelMaster em = new excelMaster();
        getDate gd = new getDate();
        dateConvert dc = new dateConvert();
        //Variables
        //string tableName = "[ShipmentData]";
        string osdate = DateTime.Now.ToString("yyyyMMdd");
        //Date is acquired in class below, using filename
        //string header = "ShipDate, GLDate, ItemIdentifer, ItemUPCCode, ItemDescription, eCommItemIdentifier, Retailer, WarehouseIdentifier, ShipToAddress, ShipToCity, ShipToState, ShipToPostalCode, QuantityShipped, WholesaleDollarsShipped";
        //int[] excelNums = new int[2] { 13, 14};
        //int startingNum = 14;
        string[] master = new string[300000];
        int mcount = 0;

        public void excelGenStart(string tableName, string header, int[] excelNums, int[] dateNums, int startingNum)
        {
            //Get Shipment Files
            string[] files = rf.ProcessDirectory(Directory.GetCurrentDirectory());
            files = fc.filterNoName(files, ".csv", ".xlsx");
            string path = files[0];
            string date = DateTime.Now.ToString("MM-dd-yyyy");
            //Process Excel Doc
            object[,] excelValues = em.excelMasterStart(path);
            //Setup loop for insert querys
            excelValues = em.filterExcel(excelValues, excelNums, startingNum);
            for (int i = startingNum; i < em.rowCountF + 1; i++)
            {
                string excelInsert = $"Insert into {tableName} ({header}) VALUES (";
                string valueHolder = "";
                //Convert Columns to dates
                for(int index = 1;index < dateNums.Length+1;index++)
                {
                    excelValues[i, index] = dc.dateConvertStart(excelValues[i, index].ToString(), "Date");
                }
                for (int index = 1; index < em.colCountF + 1; index++)
                {
                        
                    valueHolder = valueHolder + excelValues[i, index].ToString();
                    //If it is not the last item in the column, add a comma
                    if (index != em.colCountF)
                        valueHolder = valueHolder + ", ";
                }
                //combine all parts
                excelInsert = $"{excelInsert} {valueHolder})";
                Console.WriteLine(excelInsert);
                //add to query array
                master[mcount] = excelInsert;
                mcount++;

                
            }
            master = master.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            System.IO.File.WriteAllLines(Directory.GetCurrentDirectory() + $@"\{osdate}_EXCELTOSQL_output.txt", master);

        }
    }
}
