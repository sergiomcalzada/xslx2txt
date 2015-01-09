using System;

using CommandLine;
using CommandLine.Text;

namespace xslx2textconsole
{
    public class ArgumentOptions
    {

        [Option('d', "directory", Required = false, HelpText = "Directory to scan.")]
        public string Directory { get; set; }

        [Option('p', "pattern", Required = false, HelpText = "Pattner to scan.", DefaultValue = "*.xlsx")]
        public string Pattern { get; set; }

        [Option('f', "folder-only", Required = false, HelpText = "Scan subfolders")]
        public bool FolderOnly { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this, current => HelpText.DefaultParsingErrorsHandler(this, current));
        }

        public ArgumentOptions()
        {
            this.Directory = AppDomain.CurrentDomain.BaseDirectory;
        }
    }
}