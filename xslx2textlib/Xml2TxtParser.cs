using System.IO;
using System.Text;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace xslx2textlib
{
    public class Xml2TxtParser
    {
        public string Parse(string source)
        {
            var sourceFile = System.IO.File.OpenRead(source);
            var targetFile = this.GetTargetFile(source);
            Transform(sourceFile, targetFile);
            return targetFile.Name;
        }


        private FileStream GetTargetFile(string source)
        {
            var targetFileName = Path.ChangeExtension(source, "txt");
            return System.IO.File.Create(targetFileName);
        }


        private void Transform(FileStream sourceFile, FileStream targetFile)
        {

            using (var writer = new StreamWriter(targetFile))
            {
                var hssfwb = WorkbookFactory.Create(sourceFile, ImportOption.SheetContentOnly);
                foreach (ISheet sheet in hssfwb)
                {
                    writer.WriteLine("[{0}]", sheet.SheetName);
                    for (int r = 0; r <= sheet.LastRowNum; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        if (row != null) //null is when the row only contains empty cells 
                        {

                            foreach (ICell cell in row)
                            {
                                writer.Write("{0}\t", cell);
                            }
                        }
                        writer.WriteLine();
                    }
                }


            }
        }
    }
}