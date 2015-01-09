using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace xslx2textlib
{
    public class FileFinder
    {
        public IEnumerable<string> FindFiles(string path, string pattern, bool recursive)
        {
            var options = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            if (!Directory.Exists(path))
            {
                Console.WriteLine("Invalid directory name");
                return null;
            }

            var files = Directory.GetFiles(path, pattern, options);
            return files.AsEnumerable();
            
        }
    }
}
