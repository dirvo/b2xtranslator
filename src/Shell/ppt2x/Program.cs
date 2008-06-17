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
using System.Xml;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;
using System.Globalization;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.PresentationMLMapping;

namespace DIaLOGIKa.b2xtranslator.ppt2x
{
    public class Program
    {
        private static string inputFile;
        private static string outputFile;

        public static void Main(string[] args)
        {
            inputFile = args[0];
            outputFile = args.Length > 1 ? args[1] : null;

            //copy processing file
            ProcessingFile procFile = new ProcessingFile(inputFile);

            //make output file name
            if (outputFile == null)
            {
                if (inputFile.Contains("."))
                {
                    outputFile = inputFile.Remove(inputFile.LastIndexOf(".")) + ".pptx";
                }
                else
                {
                    outputFile = inputFile + ".pptx";
                }
            }

            //start time
            DateTime start = DateTime.Now;

            //open the reader
            using (StructuredStorageFile reader = new StructuredStorageFile(procFile.File.FullName))
            {

                //parse the document
                PowerpointDocument ppt = new PowerpointDocument(reader);

                using (PresentationDocument pptx = PresentationDocument.Create(outputFile))
                {
                    // Setup the writer
                    XmlWriterSettings xws = new XmlWriterSettings();
                    xws.OmitXmlDeclaration = false;
                    xws.CloseOutput = true;
                    xws.Encoding = Encoding.UTF8;
                    xws.ConformanceLevel = ConformanceLevel.Document;

                    // Setup the context
                    ConversionContext context = new ConversionContext(ppt);
                    context.WriterSettings = xws;
                    context.Pptx = pptx;

                    // Write presentation.xml
                    ppt.Convert(new PresentationPartMapping(context));
                }

                DateTime end = DateTime.Now;
                TimeSpan diff = end.Subtract(start);
                TraceLogger.Info("Conversion of file {0} finished in {1} seconds", inputFile, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}