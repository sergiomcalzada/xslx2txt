﻿using System;
using System.Diagnostics;

using xslx2textlib;

namespace xslx2textconsole
{

    class Program
    {



        static void Main(string[] args)
        {
            var finder = new FileFinder();
            var parser = new Xml2TxtParser();

            var files = finder.FindFiles(AppDomain.CurrentDomain.BaseDirectory, "*.xlsx", true);
            var fileWatcher = new Stopwatch();
            var mainWatcher = new Stopwatch();
            mainWatcher.Start();
            foreach (var file in files)
            {
                Console.Write("Transforming {0} started. ", file);
                fileWatcher.Restart();
                try
                {
                    parser.Parse(file);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    Console.WriteLine("Press any key to continue.");
                    Console.ReadKey();
                }
                Console.WriteLine("End in {0}s\n", fileWatcher.Elapsed.Seconds);
            }
            Console.WriteLine("Total elapsed time {0}m", mainWatcher.Elapsed.TotalMinutes);
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
