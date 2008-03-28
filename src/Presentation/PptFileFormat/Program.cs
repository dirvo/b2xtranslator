using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Utils;
using PptFileFormat.Records;
using System.IO;

namespace PptFileFormat
{
    class Program
    {
        static void Main(string[] args)
        {
            const string outputDir = "dumps";

            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);
            
            Directory.CreateDirectory(outputDir);

            string inputFile = args[0];
            ProcessingFile procFile = new ProcessingFile(inputFile);
            
            //open the reader
            StorageReader reader = new StorageReader(procFile.File.FullName);

            foreach (DirectoryEntry entry in reader.AllStreamEntries)
                System.Console.WriteLine(entry.Path);

            System.Console.WriteLine();

            PowerpointDocument pptDoc = new PowerpointDocument(reader);

            foreach (Record record in pptDoc)
            {
                UnknownRecord unknownRecord = record as UnknownRecord;

                if (unknownRecord != null)
                {
                    string filename = String.Format(@"{0}\{1}.record", outputDir, unknownRecord.GetIdentifier());

                    using (FileStream fs = new FileStream(filename, FileMode.Create))
                    {
                        unknownRecord.DumpToStream(fs);
                    }
                }
            }

            System.Console.WriteLine(pptDoc);

            // Let's make development as easy as pie.
            System.Diagnostics.Debugger.Break();
        }
    }
}

