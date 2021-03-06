﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelToSQL
{
    class Program
    {
        static void Main(string[] args)
        {
            //References to Classes
            generateConfig cfg = new generateConfig();
            readFiles rf = new readFiles();
            filterClass fil = new filterClass();
            processCfg pCfg = new processCfg();
            excelGenerate e = new excelGenerate();
            //Check if config file exists in current dir
            string[] files = rf.ProcessDirectory(Directory.GetCurrentDirectory());
            files = fil.filter(files, ".txt", ".txt","cfg");
            //if any file contains name: cfg then cfg exists, process config
            if (files.Any(s => s.Contains("cfg")))
            {
                string[] values = pCfg.processCfgStart(files[0]);
                for (int i = 0; i < values.Length; i++)
                {
                    Console.WriteLine(values[i]);
                }
                Console.WriteLine(pCfg.startingNumF);
                for (int i = 0; i < pCfg.numValuesF.Length; i++)
                {
                    Console.WriteLine(pCfg.numValuesF[i]);
                }
                for (int i = 0; i < pCfg.dateValuesF.Length; i++)
                {
                    Console.WriteLine(pCfg.dateValuesF[i]);
                }

                e.excelGenStart(values[0], values[1], pCfg.numValuesF, pCfg.dateValuesF, pCfg.startingNumF);
                Console.WriteLine("Finished writing querys");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();

            }
            else
            {
                //Preform Initial Config setup
                cfg.genConfigStart();
            }
        }
    }
}
