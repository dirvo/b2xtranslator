using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.Utils
{
    public class ProcessingFile
    {
        public FileInfo File;

        public ProcessingFile(string inputFile)
        {
            FileInfo inFile = new FileInfo(inputFile);

            this.File = inFile.CopyTo(System.IO.Path.GetTempFileName(), true);
            this.File.IsReadOnly = false;
        }
    }
}
