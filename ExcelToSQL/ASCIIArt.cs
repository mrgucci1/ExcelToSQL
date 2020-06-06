using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSQL
{

    class ASCIIArt
    {
        public void printLogo()
        {
            var arr = new[]
            {
                @"  _____  _____ ___ _      _____ ___     ___  ___  _   ",
                @" | __\ \/ / __| __| |    |_   _/ _ \   / __|/ _ \| |  ",
                @" | _| >  < (__| _|| |__    | || (_) |  \__ \ (_) | |__",
                @" |___/_/\_\___|___|____|   |_| \___/   |___/\__\_\____|",
                @"                                                      "
            };
            Console.WindowWidth = 160;
            Console.WriteLine("\n");
            Console.ForegroundColor = ConsoleColor.Green;
            foreach (string line in arr)
                Console.WriteLine(line);
            Console.WriteLine("\n");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Created by Keaton Dalquist");
            Console.WriteLine("\n");
            Console.ResetColor();
        }

    }
}
