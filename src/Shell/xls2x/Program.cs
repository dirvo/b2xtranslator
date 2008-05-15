/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using System.Xml;
using System.IO;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;
using System.Globalization;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.Spreadsheet;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;

namespace xls2x
{
    class Program
    {
        // Some static variables to store input data 
        private static string inputFileString;
        private static string outputDir; 

        static void Main(string[] args)
        {
            Program.parseArgs(args);

            if (Directory.Exists(Program.outputDir))
                Directory.Delete(Program.outputDir);

            Directory.CreateDirectory(Program.outputDir);
            Console.Out.WriteLine("Files will be created into Directory {0}", Program.outputDir); 

            
            try
            {
                //copy processing file
                // ProcessingFile procFile = new ProcessingFile(inputFileString);

                //open the reader
                using (StructuredStorageFile reader = new StructuredStorageFile(inputFileString))
                {

                    //parse the document
                    XlsDocument xlsDoc = new XlsDocument(reader);
                    using (SpreadsheetDocument spreadx = SpreadsheetDocument.Create("testfile.xlsx"))
                    {

                        //Setup the writer
                        XmlWriterSettings xws = new XmlWriterSettings();
                        xws.OmitXmlDeclaration = false;
                        xws.CloseOutput = true;
                        xws.Encoding = Encoding.UTF8;
                        xws.ConformanceLevel = ConformanceLevel.Document;

                        ExcelContext xlsContext = new ExcelContext(xlsDoc, xws);
                        xlsContext.SpreadDoc = spreadx; 
                        
                        // Converts the sst data !!!
                        xlsDoc.workBookData.SstData.Convert(new SSTMapping(xlsContext));
                        foreach (BoundSheetData var in xlsDoc.workBookData.boundSheetDataList)
                        {
                            var.Convert(new WorksheetMapping(xlsContext)); 
                        }
                        xlsDoc.workBookData.Convert(new WorkbookMapping(xlsContext)); 
                        
                    }
                    reader.Close(); 
                }

            }
            catch (Exception ex)
            {
                Console.Out.WriteLine("Following exception occured \n {0} \n\n Stacktrace: \n\n {1}", ex.Message, ex.StackTrace); 
            }

        }


        /// <summary>
        /// Parses the arguments 
        /// </summary>
        /// <param name="args"></param>
        static void parseArgs(String[] args)
        {
            Program.inputFileString = args[0]; 
            Program.outputDir = args[1]; 
        }

    }
}
