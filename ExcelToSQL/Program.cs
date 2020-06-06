using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSQL
{
    class Program
    {
        static void Main(string[] args)
        {
            //References to Classes
            generateConfig cfg = new generateConfig();
            //Preform Initial Config setup
            cfg.genConfigStart();

        }
    }
}
