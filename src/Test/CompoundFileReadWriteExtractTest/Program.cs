using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Writer;

namespace CompoundFileReadWriteExtractTest
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
                    Dictionary<string, KeyValuePair<string,Guid> > pathNames1 = new Dictionary<string, KeyValuePair<string,Guid> >();
                    foreach (DirectoryEntry entry in streamEntries)
                    {
                        string name = entry.Path;
                        for (int i = 0; i < invalidChars.Length; i++)
                        {
                            name = name.Replace(invalidChars[i], '_');
                        }
                        pathNames1.Add(entry.Path, new KeyValuePair<string,Guid>(name,entry.ClsId) );
                    }
                      
                    // Create Directory Structure
                    StructuredStorageWriter sso = new StructuredStorageWriter();
                    sso.RootDirectoryEntry.setClsId(storageReader.RootDirectoryEntry.ClsId);
                    foreach (string key in pathNames1.Keys)
                    {
                        StorageDirectoryEntry sde = sso.RootDirectoryEntry;
                        string[] storages = key.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < storages.Length; i++)
                        {
                            if (i != storages.Length - 1)
                            {
                                StorageDirectoryEntry result = sde.AddStorageDirectoryEntry(storages[i]);
                                sde = (result == null) ? sde : result;
                                sde.setClsId(pathNames1[key].Value);
                                continue;
                            }
                            VirtualStream vstream = storageReader.GetStream(key);
                            sde.AddStreamDirectoryEntry(storages[i], vstream);
                        }
                    }

                    // Write sso to stream
                    MemoryStream myStream = new MemoryStream();                    
                    sso.write(myStream);

                    // Close input storage
                    storageReader.Close();
                    storageReader = null;

                    // Write stream to file

                    byte[] array = new byte[bytesToReadAtOnce];
                    int bytesRead;

                    
                    string outputFileName = Path.GetFileNameWithoutExtension(file) + "_output" + Path.GetExtension(file);
                    string path = Path.GetDirectoryName(Path.GetFullPath(file));
                    outputFileName = path + "\\" + outputFileName;

                    FileStream outputFile = new FileStream(outputFileName, FileMode.Create, FileAccess.Write);
                    myStream.Seek(0, SeekOrigin.Begin);
                    do
                    {
                        bytesRead = myStream.Read(array, 0, bytesToReadAtOnce);
                        outputFile.Write(array, 0, bytesToReadAtOnce);
                    } while (bytesRead == array.Length);
                    outputFile.Close();


                    // --------- extract streams from written file



                    storageReader = new StructuredStorageReader(outputFileName);

                    // read stream _entries
                    streamEntries = storageReader.AllStreamEntries;

                    // create valid path names
                    Dictionary<string, string> pathNames2 = new Dictionary<string, string>();
                    foreach (DirectoryEntry entry in streamEntries)
                    {
                        string name = entry.Path;
                        for (int i = 0; i < invalidChars.Length; i++)
                        {
                            name = name.Replace(invalidChars[i], '_');
                        }
                        pathNames2.Add(entry.Path, name);
                    }

                    // create output directory
                    string outputDir = '_' + (Path.GetFileName(outputFileName)).Replace('.', '_');
                    outputDir = outputDir.Replace(':', '_');
                    outputDir = Path.GetDirectoryName(outputFileName) + "\\" + outputDir;
                    Directory.CreateDirectory(outputDir);

                    // for each stream                    
                    foreach (string key in pathNames2.Keys)
                    {
                        // get virtual stream by path name
                        IStreamReader streamReader = new VirtualStreamReader(storageReader.GetStream(key));

                        // read bytes from stream, write them back to disk
                        FileStream fs = new FileStream(outputDir + "\\" + pathNames2[key] + ".stream", FileMode.Create);
                        BinaryWriter writer = new BinaryWriter(fs);
                        array = new byte[bytesToReadAtOnce];                        
                        do
                        {
                            bytesRead = streamReader.Read(array);
                            writer.Write(array, 0, bytesRead);
                            writer.Flush();
                        } while (bytesRead == array.Length);

                        writer.Close();
                        fs.Close();
                    }

                    // close storage
                    storageReader.Close();
                    storageReader = null;

                    extractionTime = DateTime.Now - begin;
                    Console.WriteLine("Streams extracted in " + String.Format("{0:N2}", extractionTime.TotalSeconds) + "s. (File: " + file + ")");
                }
                catch (Exception e)
                {
                    Console.WriteLine("*** Error: " + e.Message + " (File: " + file + ")");
                    Console.WriteLine("*** StackTrace: " + e.StackTrace + " (File: " + file + ")");
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
