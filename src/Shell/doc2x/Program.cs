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
using DIaLOGIKa.b2xtranslator.WordprocessingML;
using DIaLOGIKa.b2xtranslator.WordprocessingMLMapping;
using System.IO;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using DIaLOGIKa.b2xtranslator.Utils;

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

            //check input file
            FileInfo fi = new FileInfo(inputFile);
            if (!fi.Exists)
            {
                if (verboseLvl > VerboseLevel.None)
                {
                    Console.WriteLine("The selected input file does not exist.");
                }
                Environment.Exit(1);
            }

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

            #region conversion
            try
            {               
                //start time
                DateTime start = DateTime.Now;

                //open the reader
                StorageReader reader = new StorageReader(procFile.File.FullName);

                //parse the document
                WordDocument doc = new WordDocument(reader);

                if (!doc.FIB.fComplex)
                {
                    using (WordprocessingDocument docx = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document))
                    {
                        XmlWriterSettings xws = new XmlWriterSettings();
                        xws.OmitXmlDeclaration = false;
                        xws.CloseOutput = true;
                        xws.Encoding = Encoding.UTF8;
                        //xws.Indent = true;
                        xws.ConformanceLevel = ConformanceLevel.Document;

                        XmlWriter writer = null;

                        //Write Styles.xml
                        writer = XmlWriter.Create(docx.MainDocumentPart.StyleDefinitionsPart.GetStream(), xws);
                        doc.Styles.Convert(new StyleSheetMapping(writer, doc));
                        writer.Flush();

                        //Write Document.xml
                        writer = XmlWriter.Create(docx.MainDocumentPart.GetStream(), xws);
                        doc.Convert(new DocumentMapping(writer));
                        writer.Flush();

                        DateTime end = DateTime.Now;
                        TimeSpan diff = end.Subtract(start);
                        if (verboseLvl > VerboseLevel.Warning)
                        {
                            Console.WriteLine("Conversion finished in " + diff.TotalSeconds + " seconds");
                        }
                    }
                }
                else if (verboseLvl > VerboseLevel.None)
                {
                    Console.WriteLine(inputFile + " has been fast-saved. This format is currently not supported.");
                }
            }
            catch (ZipCreationException)
            {
                if (verboseLvl > VerboseLevel.None)
                {
                    Console.WriteLine("Could not create the outputfile.");
                    Console.WriteLine("Perhaps the specified outputfile was a directory or contained invalid characters.");
                }
            }
            catch (Exception e)
            {
                if (verboseLvl > VerboseLevel.None)
                    Console.WriteLine(e.ToString());
            }
            #endregion
        }

        /// <summary>
        /// Parses the arguments of the tool
        /// </summary>
        /// <param name="args">The args array</param>
        private static void parseArgs(string[] args)
        {
            try
            {
                inputFile = args[0];              

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
                Console.WriteLine("At least one of the required arguments was not set or was used wrongly\n");
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
            usage.AppendLine("-o: change output filename");
            usage.AppendLine("-v: verbose level");
            usage.AppendLine("\tnone (0): prints nothing");
            usage.AppendLine("\terror (1): prints all errors (default)");
            usage.AppendLine("\twarning (2): prints all errors and warnings");
            usage.AppendLine("\tinfo (3): prints all errors, warnings and infos");
            usage.AppendLine("\tdebug (4): prints all errors, warnings, infos and debug informations");
            usage.AppendLine("-?: prints this help");
            Console.WriteLine(usage.ToString());
        }
    }
}
