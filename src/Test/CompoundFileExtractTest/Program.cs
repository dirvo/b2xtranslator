/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */


using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.IO;
using System.Collections.ObjectModel;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Utils;

namespace CompoundFileExtractTest
{


    /// <summary>
    /// Test application which extracts streams from compound files.
    /// Author: math
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {            
            const int bytesToReadAtOnce = 1024;
            char[] invalidChars = Path.GetInvalidFileNameChars();

            TraceLogger.LogLevel = TraceLogger.LoggingLevel.ERROR;
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

                StorageReader storageReader = null;
                DateTime begin = DateTime.Now;
                TimeSpan extractionTime = new TimeSpan();

                try
                {                   
                    // init StorageReader
                    storageReader = new StorageReader(file);

                    // read stream entries
                    ReadOnlyCollection<DirectoryEntry> streamEntries = storageReader.AllStreamEntries;

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
                    string outputDir = file.Replace('.', '_');
                    Directory.CreateDirectory(outputDir);

                    // for each stream                    
                    foreach (string key in PathNames.Keys)
                    {
                        // get virtual stream by path name
                        VirtualStream stream = storageReader.GetStream(key);

                        // read bytes from stream, write them back to disk
                        FileStream fs = new FileStream(outputDir + "\\" + PathNames[key] + ".stream", FileMode.Create);
                        BinaryWriter writer = new BinaryWriter(fs);
                        byte[] array = new byte[bytesToReadAtOnce];
                        int bytesRead;
                        do
                        {
                            bytesRead = stream.Read(array);
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
