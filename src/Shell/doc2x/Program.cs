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
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.WordprocessingMLMapping;
using System.IO;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;

namespace DIaLOGIKa.b2xtranslator.doc2x
{
    public class Program
    {
        private static string inputFile;
        private static string outputFile;
        
        public static void Main(string[] args)
        {
            //parse arguments
            parseArgs(args);

            // let the Console listen to the Trace messages
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));

            //welcome message
            printWelcome();

            try
            {
                //copy processing file
                ProcessingFile procFile = new ProcessingFile(inputFile);

                //make output file name
                if (outputFile == null)
                {
                    if (inputFile.Contains("."))
                    {
                        outputFile = inputFile.Remove(inputFile.LastIndexOf(".")) + ".docx";
                    }
                    else
                    {
                        outputFile = inputFile + ".docx";
                    }
                }

                TraceLogger.Info("Converting file {0} into {1}", inputFile, outputFile);

                //start time
                DateTime start = DateTime.Now;

                //open the reader
                using (StructuredStorageFile reader = new StructuredStorageFile(procFile.File.FullName))
                {
                    //parse the document
                    WordDocument doc = new WordDocument(reader);

                    //convert the document
                    Converter.Convert(doc, outputFile);

                    DateTime end = DateTime.Now;
                    TimeSpan diff = end.Subtract(start);
                    TraceLogger.Info("Conversion of file {0} finished in {1} seconds", inputFile, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (FileNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (ReadBytesAmountMismatchException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", inputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (MagicNumberException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", inputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (UnspportedFileVersionException ex)
            {
                TraceLogger.Error("File {0} has been created with a Word version older than Word 97.", inputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (ByteParseException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", inputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (MappingException ex)
            {
                TraceLogger.Error("There was an error while converting file {0}: {1}", inputFile, ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (ZipCreationException ex)
            {
                TraceLogger.Error("Could not create output file {0}.", outputFile);
                //TraceLogger.Error("Perhaps the specified outputfile was a directory or contained invalid characters.");
                TraceLogger.Debug(ex.ToString());
            }
            catch (Exception ex)
            {
                TraceLogger.Error("Conversion of file {0} failed.", inputFile);
                TraceLogger.Debug(ex.ToString());
            }
        }

        /// <summary>
        /// Parses the arguments of the tool
        /// </summary>
        /// <param name="args">The args array</param>
        private static void parseArgs(string[] args)
        {
            try
            {
                if (args[0] == "-?")
                {
                    printUsage();
                    Environment.Exit(0);
                }
                else
                {
                    inputFile = args[0];
                }

                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].ToLower() == "-v")
                    {
                        //parse verbose level
                        string verbose = args[i + 1].ToLower();
                        int vLvl;
                        if (Int32.TryParse(verbose, out vLvl))
                        {
                            TraceLogger.LogLevel = (TraceLogger.LoggingLevel)vLvl;
                        }
                        else if (verbose == "error")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Error;
                        }
                        else if (verbose == "warning")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Warning;
                        }
                        else if (verbose == "info")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Info;
                        }
                        else if (verbose == "debug")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Debug;
                        }
                        else if (verbose == "none")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.None;
                        }
                    }
                    else if (args[i].ToLower() == "-o")
                    {
                        //parse output file name
                        outputFile = args[i + 1];
                    }
                }
            }
            catch (Exception)
            {
                TraceLogger.Error("At least one of the required arguments was not correctly set.\n");
                printUsage();
                Environment.Exit(1);
            }
        }



        /// <summary>
        /// Prints the usage of the tool
        /// </summary>
        private static void printUsage()
        {
            StringBuilder usage = new StringBuilder();
            usage.AppendLine("Usage: doc2x filename [-o filename] [-v level] [-?]");
            usage.AppendLine("-o <filename>  change output filename");
            usage.AppendLine("-v <level>     set trace level, where <level> is one of the following:");
            usage.AppendLine("               none (0)    print nothing");
            usage.AppendLine("               error (1)   print all errors");
            usage.AppendLine("               warning (2) print all errors and warnings");
            usage.AppendLine("               info (3)    print all errors, warnings and infos (default)");
            usage.AppendLine("               debug (4)   print all errors, warnings, infos and debug messages");
            usage.AppendLine("-?             print this help");
            Console.WriteLine(usage.ToString());
        }


        /// <summary>
        /// Prints the heading row of the tool
        /// </summary>
        private static void printWelcome()
        {
            bool backup = TraceLogger.EnableTimeStamp;
            TraceLogger.EnableTimeStamp = false;
            StringBuilder welcome = new StringBuilder();
            welcome.Append("Welcome to doc2x.exe (r");
            welcome.Append(getRevision());
            welcome.Append(")\n");
            welcome.Append("Copyright (c) 2008, DIaLOGIKa. All rights reserved.");
            welcome.Append("\n");
            TraceLogger.Simple(welcome.ToString());
            TraceLogger.EnableTimeStamp = backup;
        }


        /// <summary>
        /// Returns the revision that is stored in the embedded resource "revision.txt".
        /// Returns -1 if something goes wrong
        /// </summary>
        /// <returns></returns>
        private static int getRevision()
        {
            int rev = -1;

            try
            {
                Assembly a = Assembly.GetExecutingAssembly();
                Stream s = a.GetManifestResourceStream("DIaLOGIKa.b2xtranslator.doc2x.revision.txt");
                StreamReader reader = new StreamReader(s);
                rev = Int32.Parse(reader.ReadLine());
                s.Close();
            }
            catch (Exception) { }

            return rev;
        }
    }
}
