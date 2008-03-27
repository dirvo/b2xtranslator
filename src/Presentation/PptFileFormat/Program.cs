using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Utils;

namespace PptFileFormat
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = args[0];
            ProcessingFile procFile = new ProcessingFile(inputFile);

            //open the reader
            StorageReader reader = new StorageReader(procFile.File.FullName);

            foreach (DirectoryEntry entry in reader.AllStreamEntries)
                System.Console.WriteLine(entry.Path);

            System.Console.WriteLine();

            PowerpointDocument pptDoc = new PowerpointDocument(reader);

            System.Console.ReadLine();
        }
    }
}
