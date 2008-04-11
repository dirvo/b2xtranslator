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
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.IO;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.WordprocessingMLMapping;

namespace DocTranslatorTest
{
    class Program
    {
        private static string file;

        static void Main(string[] args)
        {
            //parse arguments
            parseArgs(args);

            //start time
            DateTime start = DateTime.Now;

            //open the reader
            StorageReader reader = new StorageReader(file);

            //parse the document
            WordDocument doc = new WordDocument(reader);
            
            //starting
            if (!doc.FIB.fComplex)
            {
                using (WordprocessingDocument docx = WordprocessingDocument.Create(file + "x", WordprocessingDocumentType.Document))
                {
                    XmlWriterSettings xws = new XmlWriterSettings();
                    xws.OmitXmlDeclaration = false;
                    xws.CloseOutput = true;
                    xws.Encoding = Encoding.UTF8;
                    xws.ConformanceLevel = ConformanceLevel.Document;

                    //write docProps/app.xml
                    //doc.DocumentProperties.Convert(new ApplicationPropertiesMapping(docx.AddAppPropertiesPart(), xws));

                    //write settings.xml
                    doc.DocumentProperties.Convert(new SettingsMapping(docx.MainDocumentPart.AddSettingsPart(), xws, doc.FIB));

                    //Write styles.xml
                    doc.Styles.Convert(new StyleSheetMapping(docx.MainDocumentPart.AddStyleDefinitionsPart(), xws, doc));

                    //Write numbering.xml
                    doc.ListTable.Convert(new NumberingMapping(docx.MainDocumentPart.AddNumberingDefinitionsPart(), xws, doc));

                    //Write fontTable.xml
                    doc.FontTable.Convert(new FontTableMapping(docx.MainDocumentPart.AddFontTablePart(), xws));

                    //Write document.xml
                    doc.Convert(new DocumentMapping(docx.MainDocumentPart, xws));

                    DateTime end = DateTime.Now;
                    TimeSpan diff = end.Subtract(start);
                    Console.WriteLine("Conversion finished in "  + diff.TotalSeconds + " seconds");
                }
            }
            else
            {
                Console.WriteLine(file + " has been fast-saved. This format is currently not supported.");
            }

            reader.Close();
        }

        public static void AddContent(XmlWriter writer)
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("styles", OpenXmlNamespaces.WordprocessingML);

            writer.WriteEndElement();
            writer.WriteEndDocument();

            writer.Flush();
        }

        /// <summary>
        /// Parses the arguments
        /// </summary>
        /// <param name="args"></param>
        private static void parseArgs(string[] args)
        {
            try
            {
                file = args[0];
                //FileInfo fi = new FileInfo(file);
                //method = args[1];
            }
            catch (Exception)
            {
                throw new ArgumentException();
            }
        }

        /// <summary>
        /// Prints the usage of the tool
        /// </summary>
        private static void printUsage()
        {
            Console.WriteLine("Usage: DocTranslatorTest.exe {filename}");
        }
    }
}
