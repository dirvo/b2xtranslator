using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Writer;

namespace CompoundFileWritingTest
{
    class Program
    {
        static void Main(string[] args)
        {
            const int bytesToReadAtOnce = 512;
            char[] invalidChars = Path.GetInvalidFileNameChars();

            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Error;
            ConsoleTraceListener consoleTracer = new ConsoleTraceListener();
            Trace.Listeners.Add(consoleTracer);
            Trace.AutoFlush = true;

            if (args.Length < 1)
            {
                Console.WriteLine("No parameter found. Please specify one or more compound document file(s).");
                return;
            }

            foreach (string file in args)
            {

                StructuredStorageReader storageReader = null;
                DateTime begin = DateTime.Now;
                TimeSpan extractionTime = new TimeSpan();

                try
                {
                    // init StorageReader
                    storageReader = new StructuredStorageReader(file);

                    // read stream _entries
                    ICollection<DirectoryEntry> streamEntries = storageReader.AllStreamEntries;

                    // create valid path names
                    Dictionary<string, string> PathNames = new Dictionary<string, string>();
                    foreach (DirectoryEntry entry in streamEntries)
                    {
                        string name = entry.Path;
                        for (int i = 0; i < invalidChars.Length; i++)
                        {
                            name = name.Replace(invalidChars[i], '_');
                        }
                        PathNames.Add(entry.Path, name);
                    }

                    // create output directory
                    string outputDir = '_' + file.Replace('.', '_');
                    outputDir = outputDir.Replace(':', '_');
                    Directory.CreateDirectory(outputDir);

                    //ICollection<DirectoryEntry> allEntries = storageReader.AllStreamEntries;


                    //foreach (DirectoryEntry entry in allEntries)
                    //{
                    //    if (entry.Type == DirectoryEntryType.STGTY_STREAM)
                    //    {
                    //        sso.RootDirectoryEntry.AddStreamDirectoryEntry(entry.Name);
                    //    }
                    //    else if (entry.Type == DirectoryEntryType.STGTY_STORAGE)
                    //    {
                    //        sso.RootDirectoryEntry.AddStorageDirectoryEntry(entry.Name);
                    //    }
                    //}


                    // for each stream       

                    MemoryStream tmp = new MemoryStream();
                    StructuredStorageWriter sso = new StructuredStorageWriter();
                    foreach (string key in PathNames.Keys)
                    {
                        StorageDirectoryEntry sde = sso.RootDirectoryEntry;
                        string[] storages = key.Split(new char[] {'\\'}, StringSplitOptions.RemoveEmptyEntries );
                        for (int i = 0; i < storages.Length; i++)
                        {
                            if (i != storages.Length - 1)
                            {
                                StorageDirectoryEntry result = sde.AddStorageDirectoryEntry(storages[i]);
                                sde = (result == null) ? sde : result;
                                continue;
                            }
                            VirtualStream vstream = storageReader.GetStream(key);
                            sde.AddStreamDirectoryEntry(storages[i], vstream);
                        }                        
                        

                        // get virtual stream by path name
                        
                        //sso.RootDirectoryEntry.AddStorageDirectoryEntry("ABCDEFG");
                        //sso.RootDirectoryEntry.AddStorageDirectoryEntry("VWXYZ");
                        //sso.RootDirectoryEntry.AddStorageDirectoryEntry("ABCFG");
                        //sso.RootDirectoryEntry.AddStorageDirectoryEntry("ABCDE");
                        //// read bytes from stream, write them back to disk
                        //FileStream fs = new FileStream(outputDir + "\\" + PathNames[key] + ".stream", FileMode.Create);
                        //BinaryWriter writer = new BinaryWriter(fs);
                        //byte[] array = new byte[bytesToReadAtOnce];
                        //int bytesRead;
                        //do
                        //{
                        //    bytesRead = streamReader.Read(array);
                        //    writer.Write(array, 0, bytesRead);
                        //    writer.Flush();
                        //} while (bytesRead == array.Length);

                        //writer.Close();
                        //fs.Close();
                    }

                    MemoryStream myStream = new MemoryStream();
                    sso.write(myStream);

                    byte[] array = new byte[bytesToReadAtOnce];
                    int bytesRead;

                    FileStream outputFile = new FileStream(Path.GetFileNameWithoutExtension(file) + "_output" + Path.GetExtension(file), FileMode.Create, FileAccess.Write);
                    
                    myStream.Seek(0, SeekOrigin.Begin);
                    do
                    {
                        bytesRead = myStream.Read(array, 0, bytesToReadAtOnce);
                        outputFile.Write(array, 0, bytesToReadAtOnce);
                    } while (bytesRead == array.Length);

                    outputFile.Close();

                    // close storage
                    storageReader.Close();
                    storageReader = null;

                    extractionTime = DateTime.Now - begin;
                    Console.WriteLine("Streams extracted in " + String.Format("{0:N2}", extractionTime.TotalSeconds) + "s. (File: " + file + ")");
                }
                catch (Exception e)
                {
                    Console.WriteLine("*** Error: " + e.Message + " (File: " + file + ")");
                }
                finally
                {
                    if (storageReader != null)
                    {
                        storageReader.Close();
                    }
                }
            }
        }
    }
}
