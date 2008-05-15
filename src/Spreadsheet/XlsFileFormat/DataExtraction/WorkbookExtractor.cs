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
using System.Collections;
using System.Text;
using System.IO;
using System.Collections.ObjectModel;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer; 

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{

    /// <summary>
    /// Extracts the workbook stream !!
    /// </summary>
    public class WorkbookExtractor : Extractor, IVisitable
    {    
        public string buffer;
        public long oldOffset; 

        public List<BOUNDSHEET> boundsheets;

        public List<BoundSheetData> sheets;

        public WorkBookData workBookData; 

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Reader</param>
        public WorkbookExtractor(VirtualStreamReader reader, WorkBookData workBookData) 
           : base(reader) 
        {         
            this.boundsheets = new List<BOUNDSHEET>();
            this.sheets = new List<BoundSheetData>();
            this.workBookData = workBookData; 
            this.oldOffset = 0; 

            this.extractData(); 
        }

        /// <summary>
        /// Extracts the data from the stream 
        /// </summary>
        public override void extractData()
        {
            BiffHeader bh;
            StreamWriter sw = null;
            sw = new StreamWriter(Console.OpenStandardOutput());
            
            try
            {
                while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
                {
                    bh.id = (RecordNumber)this.StreamReader.ReadUInt16();

                    bh.length = this.StreamReader.ReadUInt16();

                    if (bh.id == RecordNumber.BOUNDSHEET)
                    {
                        // Creates a BoundSheetData element
                        BoundSheetData bsd = new BoundSheetData();
                        this.workBookData.addBoundSheetData(bsd); 

                        // Extracts the Boundsheet data 
                        BOUNDSHEET bs = new BOUNDSHEET(this.StreamReader, bh.id, bh.length);
                        
                        this.oldOffset = this.StreamReader.BaseStream.Position;
                        this.StreamReader.BaseStream.Seek(bs.lbPlyPos, SeekOrigin.Begin);
                        bsd.worksheetName = bs.getBoundsheetName(); 
                        BoundSheetExtractor se = new BoundSheetExtractor(this.StreamReader,bsd);
                        this.StreamReader.BaseStream.Seek(oldOffset, SeekOrigin.Begin); 
                        
                        sw.Write(bs.ToString());
                    } else if (bh.id == RecordNumber.SST)
                    {
                        SST sst = new SST(this.StreamReader, bh.id, bh.length);
                        this.workBookData.SstData = new SSTData(sst); 
                        sw.Write(sst.ToString()); 
                    }

                    else if (bh.id == RecordNumber.EOF)
                    {
                        // this.StreamReader.BaseStream.Seek(0, SeekOrigin.End); 
                        sw.Write("EOF"); 
                    } 
                    else
                    {
                        // every other record which is not implemented 

                        byte[] buffer = new byte[bh.length];
                        buffer = this.StreamReader.ReadBytes(bh.length);
                        if (bh.length != buffer.Length)
                            sw.WriteLine("EOF");

                        sw.Write("BIFF {0}\t{1}\t", bh.id, bh.length);
                        //Dump(buffer);
                        int count = 0;
                        foreach (byte b in buffer)
                        {
                            sw.Write("{0:X02} ", b);
                            count++;
                            if (count % 16 == 0 && count < buffer.Length)
                                sw.Write("\n\t\t\t");
                        }
                        sw.Write("\n");
                    }

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            sw.Close();
        }

        /// <summary>
        /// A normal overload ToString Method 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string returnvalue = "Workbook";
            return returnvalue;
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WorkbookExtractor>)mapping).Apply(this);
        }

        #endregion
    }
}
