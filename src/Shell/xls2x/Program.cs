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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using System.Reflection;
using Microsoft.Win32;
using DIaLOGIKa.b2xtranslator.Shell;

namespace DIaLOGIKa.b2xtranslator.xls2x
{
    class Program : CommandLineTranslator
    {
        public static string ToolName = "xls2x";
        public static string RevisionResource = "DIaLOGIKa.b2xtranslator.xls2x.revision.txt";
        public static string ContextMenuInputExtension = ".xls";
        public static string ContextMenuText = "Convert to .xlsx";

        static void Main(string[] args)
        {
            ParseArgs(args, ToolName);

            InitializeLogger();

            PrintWelcome(ToolName, RevisionResource);

            if (CreateContextMenuEntry)
            {
                // create context menu entry
                try
                {
                    TraceLogger.Info("Creating context menu entry for xls2x ...");
                    RegisterForContextMenu(GetContextMenuKey(ContextMenuInputExtension, ContextMenuText));
                    TraceLogger.Info("Succeeded.");
                }
                catch (Exception)
                {
                    TraceLogger.Info("Failed. Sorry :(");
                }
            }
            else
            {
                try
                {
                    //copy processing file
                    ProcessingFile procFile = new ProcessingFile(InputFile);

                    //make output file name
                    if (ChoosenOutputFile == null)
                    {
                        if (InputFile.Contains("."))
                        {
                            ChoosenOutputFile = InputFile.Remove(InputFile.LastIndexOf(".")) + ".xlsx";
                        }
                        else
                        {
                            ChoosenOutputFile = InputFile + ".xlsx";
                        }
                    }

                    TraceLogger.Info("Converting file {0} into {1}", InputFile, ChoosenOutputFile);

                    //start time
                    DateTime start = DateTime.Now;

                    //parse the document
                    using (StructuredStorageReader reader = new StructuredStorageReader(procFile.File.FullName))
                    {
                        XlsDocument xlsDoc = new XlsDocument(reader);
                        using (SpreadsheetDocument spreadx = SpreadsheetDocument.Create(ChoosenOutputFile))
                        {

                            //Setup the writer
                            XmlWriterSettings xws = new XmlWriterSettings();
                            xws.CloseOutput = true;
                            xws.Encoding = Encoding.UTF8;
                            xws.ConformanceLevel = ConformanceLevel.Document;

                            ExcelContext xlsContext = new ExcelContext(xlsDoc, xws);
                            xlsContext.SpreadDoc = spreadx;

                            // Converts the sst data !!!
                            if (xlsDoc.workBookData.SstData != null)
                                xlsDoc.workBookData.SstData.Convert(new SSTMapping(xlsContext));

                            // creates the styles.xml
                            if (xlsDoc.workBookData.styleData != null)
                                xlsDoc.workBookData.styleData.Convert(new StylesMapping(xlsContext));

                            // creates the Spreadsheets
                            foreach (WorkSheetData var in xlsDoc.workBookData.boundSheetDataList)
                            {
                                if (var.boundsheetRecord.sheetType == BOUNDSHEET.sheetTypes.worksheet)
                                {
                                    var.Convert(new WorksheetMapping(xlsContext));
                                }
                            }
                            int sbdnumber = 1;
                            foreach (SupBookData sbd in xlsDoc.workBookData.supBookDataList)
                            {
                                if (!sbd.SelfRef)
                                {
                                    sbd.Number = sbdnumber;
                                    sbdnumber++;
                                    sbd.Convert(new ExternalLinkMapping(xlsContext));
                                }
                            }

                            xlsDoc.workBookData.Convert(new WorkbookMapping(xlsContext));
                        }
                        reader.Close();
                        DateTime end = DateTime.Now;
                        TimeSpan diff = end.Subtract(start);
                        TraceLogger.Info("Conversion of file {0} finished in {1} seconds", InputFile, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
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
                catch (ZipCreationException ex)
                {
                    TraceLogger.Error("Could not create output file {0}.", ChoosenOutputFile);
                    TraceLogger.Debug(ex.ToString());
                }
                catch (Exception ex)
                {
                    TraceLogger.Error("Conversion of file {0} failed.", InputFile);
                    TraceLogger.Debug(ex.ToString());
                }
            }
        }
    }
}
