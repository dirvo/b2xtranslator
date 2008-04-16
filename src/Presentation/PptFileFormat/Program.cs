using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Tools;
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
            StructuredStorageFile reader = new StructuredStorageFile(procFile.File.FullName);

            foreach (DirectoryEntry entry in reader.AllStreamEntries)
                Console.WriteLine(entry.Path);

            Console.WriteLine();

            PowerpointDocument pptDoc = new PowerpointDocument(reader);

            // Dump unknown records
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

            // Output record tree
            Console.WriteLine(pptDoc);
            Console.WriteLine();

            // Output text for each slide
            int slideNo = 0;
            foreach (Record record in pptDoc)
            {
                Slide slide = record as Slide;

                if (slide != null)
                {
                    slideNo++;
                    Console.WriteLine("Text for slide #{0}:", slideNo);

                    bool textFound = false;
                    foreach (Record trecord in slide)
                    {
                        TextAtom text = trecord as TextAtom;

                        if (text != null)
                        {
                            Console.WriteLine("  * {0}", text.Text);
                            textFound = true;
                        }
                    }

                    if (!textFound)
                        Console.WriteLine("  No text found");

                    Console.WriteLine();
                }
            }

            // Let's make development as easy as pie.
            System.Diagnostics.Debugger.Break();
        }
    }
}

