using System;
using System.Diagnostics;

using xslx2textlib;

namespace xslx2textconsole
{

    class Program
    {
        static void Main(string[] args)
        {
            var argumentOptions = new ArgumentOptions();
            var result = CommandLine.Parser.Default.ParseArguments(args, argumentOptions);
            if (result)
            {

                var finder = new FileFinder();
                var parser = new Xml2TxtParser();

                var files = finder.FindFiles(argumentOptions.Directory, argumentOptions.Pattern, !argumentOptions.FolderOnly);
                if (files != null)
                {
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
                            ConsoleWait("Press any key to continue.");
                        }
                        Console.WriteLine("End in {0}s\n", fileWatcher.Elapsed.Seconds);
                    }
                    Console.WriteLine("Total elapsed time {0:N}m", mainWatcher.Elapsed.TotalMinutes);
                }

                ConsoleWait();
            }
            else
            {
                Console.WriteLine(argumentOptions.GetUsage());
                ConsoleWait();
            }
        }

        private static void ConsoleWait(string msg = "Press any key to exit.")
        {
            Console.WriteLine(msg);
            Console.ReadKey();
        }
    }
}
