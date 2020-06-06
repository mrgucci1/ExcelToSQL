using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelToSQL
{
    class generateConfig
    {
        //References to classes
        ASCIIArt art = new ASCIIArt();
        public void genConfigStart()
        {
            //variables to hold config info
            string sqlHeader = "initial";
            string tableName = "initial";
            int startingRow = 2;
            int[] excelNums = new int[200];
            int[] dateNums = new int[200];
            int numCounter = 0;
            int dateCounter = 0;
            art.printLogo();
            Console.WriteLine($"Welcome to first time setup\n\nThe Purpose of this program is to turn an excel document into insertable SQL querys.\n\n");
            Console.WriteLine($"Insert the name of your SQL Table, exactly as it appears in your SQL database (Example: [CustomerData] or CustomerData. If space is in table name, add brackets. Example: [Customer Data])");
            tableName = Console.ReadLine().ToString().Trim();
            Console.WriteLine("Insert the Column Headers, Sperated by a comma. If there is a space in the column name, add brackets. (Example: PhoneNumber, Street, [House Number], Name");
            sqlHeader = Console.ReadLine().ToString().Trim();
            Console.WriteLine("Insert the TOTAL number of columns that DO NOT require single quotes. Columns that DO NOT require single quotes are often numbers or decimanls, but consult your datatypes and sql formatting (Example: if I have 4 columns with numbers in my excel document, I insert 4)");
            numCounter = Convert.ToInt32(Console.ReadLine());
            excelNums = new int[numCounter];
            for(int i = 0; i< numCounter;i++)
            {
                Console.WriteLine("Insert one of the column numbers with the datatype not requiring single quotes. (Example: If my 3rd column (Column C) is a number and does not require single quotes, I would enter 3)");
                Console.WriteLine($"You will have to do this {numCounter - i} more times");
                excelNums[i] = Convert.ToInt32(Console.ReadLine());
            }
            Console.WriteLine($"What ROW does the data actually start on? (Example: Excluding the column headers, what row does the data start on. So if in Row 1 I had: Name and Row 2 I had: Jessie then I would insert row 2)");
            startingRow = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Insert the TOTAL number of columns that are dates. (Example: if I have 2 columns with dates in my excel document, I insert 2)");
            dateCounter = Convert.ToInt32(Console.ReadLine());
            dateNums = new int[dateCounter];
            for (int i = 0; i < dateCounter; i++)
            {
                Console.WriteLine("Insert one of the column numbers with the date(s) from your excel document. (Example: If my 3rd column (Column C) is a date, I would enter 3)");
                Console.WriteLine($"You will have to do this {dateCounter - i} more times");
                dateNums[i] = Convert.ToInt32(Console.ReadLine());
            }

            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(Directory.GetCurrentDirectory() + @"\cfg.txt"))
            {
                file.WriteLine("Config file for Excel To SQL - Do Not Delete");
                file.WriteLine(tableName);
                file.WriteLine(sqlHeader);
                file.WriteLine(startingRow);
                file.WriteLine(numCounter);
                foreach (int num in excelNums)
                    file.WriteLine(num);
                file.WriteLine(dateCounter);
                foreach (int num in dateNums)
                    file.WriteLine(num);
            }


        }
    }
}
