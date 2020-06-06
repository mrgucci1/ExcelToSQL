using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelToSQL
{
    class processCfg
    {
        int[] numValues = new int[200];
        int[] dateNums = new int[200];
        int startingNum = 0;
        public string[] processCfgStart(string path)
        {
            int counter = 1;
            string[] values = new string[200];

            var lines = File.ReadAllLines(path);
            //Second line is tablename
            values[0] = lines[1].ToString();
            //Third Line is Table Header
            values[1] = lines[2].ToString();
            //Forth line is starting num
            startingNum = Convert.ToInt32(lines[3]);
            //Fifth line is number of numValues
            numValues = new int[Convert.ToInt32(lines[4])];
            for (int i = 0; i < Convert.ToInt32(lines[4]) ;i++)
            {
                numValues[i] = Convert.ToInt32(lines[5 + i]);
                counter++;
            }
            //nth line is number of dateValues
            dateNums = new int[Convert.ToInt32(lines[4 + counter])];
            for (int i = 0; i < Convert.ToInt32(lines[4 + counter]); i++)
            {
                dateNums[i] = Convert.ToInt32(lines[(5+counter) + i]);
            }
            numValuesF = numValues;
            dateValuesF = dateNums;
            startingNumF = startingNum;
            values = values.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            return values;
        }

        public int[] numValuesF { get; set; }
        public int[] dateValuesF { get; set; }
        public int startingNumF { get; set; }
    }
}
