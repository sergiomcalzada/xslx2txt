using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace xslx2textlib
{
    public class FileFinder
    {
        public IEnumerable<string> FindFiles(string directory, string pattern, bool recursive)
        {
            var options = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(directory, pattern, options);
            return files.AsEnumerable();
        }
    }
}
