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

namespace DIaLOGIKa.b2xtranslator.doc2x
{
    public class Program
    {
        private enum VerboseLevel
        {
            None = 0,
            Error,
            Warning,
            Info,
            Debug
        }

        private static string inputFile;
        private static string outputFile;
        private static VerboseLevel verboseLvl = VerboseLevel.Error;

        public static void Main(string[] args)
        {
            //parse arguments
            parseArgs(args);

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

                Console.WriteLine("Converting file {0}", inputFile);

                //start time
                DateTime start = DateTime.Now;

                //open the reader
                using (StructuredStorageFile reader = new StructuredStorageFile(procFile.File.FullName))
                {

                    //parse the document
                    WordDocument doc = new WordDocument(reader);

                    if (!doc.FIB.fComplex)
                    {
                        using (WordprocessingDocument docx = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document))
                        {
                            //Setup the writer
                            XmlWriterSettings xws = new XmlWriterSettings();
                            xws.OmitXmlDeclaration = false;
                            xws.CloseOutput = true;
                            xws.Encoding = Encoding.UTF8;
                            xws.ConformanceLevel = ConformanceLevel.Document;

                            //Setup the context
                            ConversionContext context = new ConversionContext(doc);
                            context.WriterSettings = xws;
                            context.Docx = docx;

                            //Write styles.xml
                            doc.Styles.Convert(new StyleSheetMapping(context));

                            //Write numbering.xml
                            doc.ListTable.Convert(new NumberingMapping(context));

                            //Write fontTable.xml
                            doc.FontTable.Convert(new FontTableMapping(context));

                            //Write document.xml
                            doc.Convert(new MainDocumentMapping(context));

                            //write settings.xml at last because of the rsid list
                            doc.DocumentProperties.Convert(new SettingsMapping(context));
                        }
                        DateTime end = DateTime.Now;
                        TimeSpan diff = end.Subtract(start);
                        if (verboseLvl > VerboseLevel.Warning)
                        {
                            Console.WriteLine("Conversion finished in " + diff.TotalSeconds + " seconds");
                        }
                    }
                    else if (verboseLvl > VerboseLevel.None)
                    {
                        Console.WriteLine(inputFile + " has been fast-saved. This format is currently not supported.");
                    }
                }
            }
            catch (DirectoryNotFoundException)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine("The input file does not exist.");
            }
            catch (FileNotFoundException)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine("The input file does not exist.");
            }
            catch (ReadBytesAmountMismatchException)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine("The input file is not a valid .doc file.");
            }
            catch (MagicNumberException)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine("The input file is not a valid .doc file.");
            }
            catch (UnspportedFileVersionException)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine("Doc2x doesn't support file older than Word 97.");
            }
            catch (ByteParseException)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine("The input file is not a valid .doc file.");
            }
            catch (ZipCreationException)
            {
                if (verboseLvl > VerboseLevel.None)
                {
                    Console.WriteLine("Could not create the output file.");
                    Console.WriteLine("Perhaps the specified outputfile was a directory or contained invalid characters.");
                }
            }
            catch (Exception e)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine("Conversion failed.");
                if (verboseLvl > VerboseLevel.Info)
                    Console.WriteLine(e.ToString());
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
                            verboseLvl = (VerboseLevel)vLvl;
                        }
                        else if (verbose == "error")
                        {
                            verboseLvl = VerboseLevel.Error;
                        }
                        else if (verbose == "warning")
                        {
                            verboseLvl = VerboseLevel.Warning;
                        }
                        else if (verbose == "info")
                        {
                            verboseLvl = VerboseLevel.Info;
                        }
                        else if (verbose == "debug")
                        {
                            verboseLvl = VerboseLevel.Debug;
                        }
                        else if (verbose == "none")
                        {
                            verboseLvl = VerboseLevel.None;
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
                Console.WriteLine("At least one of the required arguments was not correctly set.\n");
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
            usage.AppendLine("-o <filename>   change output filename");
            usage.AppendLine("-v <level>      set trace level, where <level> is one of the following:");
            usage.AppendLine("    none (0)    print nothing");
            usage.AppendLine("    error (1)   print all errors (default)");
            usage.AppendLine("    warning (2) print all errors and warnings");
            usage.AppendLine("    info (3)    print all errors, warnings and infos");
            usage.AppendLine("    debug (4)   print all errors, warnings, infos and debug messages");
            usage.AppendLine("-?              print this help");
            Console.WriteLine(usage.ToString());
        }
    }
}
