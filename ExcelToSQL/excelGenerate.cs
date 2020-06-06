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
    class excelRDShipment
    {
        //Reference to classes
        readFiles rf = new readFiles();
        filterClass fc = new filterClass();
        excelMaster em = new excelMaster();
        getDate gd = new getDate();
        dateConvert dc = new dateConvert();
        //Variables
        string tableName = "[ShipmentData]";
        string osdate = DateTime.Now.ToString("yyyyMMdd");
        //Date is acquired in class below, using filename
        string header = "ShipDate, GLDate, ItemIdentifer, ItemUPCCode, ItemDescription, eCommItemIdentifier, Retailer, WarehouseIdentifier, ShipToAddress, ShipToCity, ShipToState, ShipToPostalCode, QuantityShipped, WholesaleDollarsShipped";
        int[] excelNums = new int[2] { 13, 14};
        int startingNum = 14;
        string[] master = new string[300000];
        int mcount = 0;

        public void excelRDStart()
        {
            //Get Shipment Files
            string[] files = rf.ProcessDirectory(Directory.GetCurrentDirectory());
            files = fc.filterNoName(files, ".csv", ".xlsx");
            for (int count = 0; count < files.Length; count++)
            {
                string[] querys = new string[10000];
                int counter = 0;
                string path = files[count];
                string date = DateTime.Now.ToString("MM-dd-yyyy");
                //Process Excel Doc
                object[,] excelValues = em.excelMasterStart(path);
                //Setup loop for insert querys
                excelValues = em.filterExcel(excelValues, excelNums, startingNum);
                for (int i = startingNum; i < em.rowCountF + 1; i++)
                {
                    string excelInsert = $"Insert into {tableName} ({header}) VALUES (";
                    string valueHolder = "";
                    excelValues[i, 1] = dc.dateConvertStart(excelValues[i, 1].ToString(), "ShipDate");
                    excelValues[i, 2] = dc.dateConvertStart(excelValues[i, 2].ToString(), "GL/Date");
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
                    querys[counter] = excelInsert;
                    master[mcount] = excelInsert;
                    mcount++;
                    counter++;
                }
                querys = querys.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                System.IO.File.WriteAllLines(Directory.GetCurrentDirectory() + $@"\{osdate}_RD_SHIPMENT_{count + 1}_output.txt", querys);
            }
            if(files.Length > 1)
            {
                master = master.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                System.IO.File.WriteAllLines(Directory.GetCurrentDirectory() + $@"\{osdate}_RD_SHIPMENT_ALL_output.txt", master);
            }
            
        }
    }
}
