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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.StyleData; 

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{

    /// <summary>
    /// Extracts the workbook stream !!
    /// </summary>
    public class WorkbookExtractor : Extractor, IVisitable
    {
        public string buffer;
        public long oldOffset;

        public List<BoundSheet8> boundsheets;
        public List<ExternSheet> externSheets;
        public List<SupBook> supBooks;
        public List<XCT> XCTList;
        public List<CRN> CRNList; 

        public WorkBookData workBookData;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Reader</param>
        public WorkbookExtractor(VirtualStreamReader reader, WorkBookData workBookData)
            : base(reader)
        {
            this.boundsheets = new List<BoundSheet8>();
            this.supBooks = new List<SupBook>(); 
            this.externSheets = new List<ExternSheet>();
            this.XCTList = new List<XCT>();
            this.CRNList = new List<CRN>(); 
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
            bool firstBOF = true; 
            try
            {
                while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
                {
                    bh.id = (RecordType)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();
                    // Debugging output 
                    TraceLogger.DebugInternal("BIFF {0}\t{1}\t", bh.id, bh.length);

                    if (bh.id == RecordType.BoundSheet8)
                    {
                        // Creates a WorkSheetData element
                        WorkSheetData bsd = new WorkSheetData();
                        this.workBookData.addBoundSheetData(bsd);

                        // Extracts the Boundsheet data 
                        BoundSheet8 bs = new BoundSheet8(this.StreamReader, bh.id, bh.length);

                        this.oldOffset = this.StreamReader.BaseStream.Position;
                        this.StreamReader.BaseStream.Seek(bs.lbPlyPos, SeekOrigin.Begin);
                        bsd.worksheetName = bs.getBoundsheetName();
                        bsd.boundsheetRecord = bs;
                        WorksheetExtractor se = new WorksheetExtractor(this.StreamReader, bsd);
                        this.StreamReader.BaseStream.Seek(oldOffset, SeekOrigin.Begin);
                        TraceLogger.DebugInternal(bs.ToString());
                    }
                    else if (bh.id == RecordType.Template)
                    {
                        this.workBookData.Template = true;
                    }
                    else if (bh.id == RecordType.SST)
                    {
                        /* reads the shared string table biff record and following continue records 
                         * creates an array of bytes and then puts that into a memory stream 
                         * this all is used to create a longer biffrecord then 8224 bytes. If theres a string 
                         * beginning in the SST that is then longer then the 8224 bytes, it continues in the 
                         * CONTINUE BiffRecord, so the parser has to read over the SST border. 
                         * The problem here is, that the parser has to overread the continue biff record header 
                        */
                        SST sst;
                        UInt16 length = bh.length;

                        // save the old offset from this record begin 
                        this.oldOffset = this.StreamReader.BaseStream.Position;
                        // create a list of bytearrays to store the following continue records 
                        // List<byte[]> byteArrayList = new List<byte[]>();
                        byte[] buffer = new byte[length];
                        LinkedList<VirtualStreamReader> vsrList = new LinkedList<VirtualStreamReader>();
                        buffer = this.StreamReader.ReadBytes((int)length);
                        // byteArrayList.Add(buffer);

                        // create a new memory stream and a new virtualstreamreader 
                        MemoryStream bufferstream = new MemoryStream(buffer);
                        VirtualStreamReader binreader = new VirtualStreamReader(bufferstream);
                        BiffHeader bh2;
                        bh2.id = (RecordType)this.StreamReader.ReadUInt16();

                        while (bh2.id == RecordType.Continue)
                        {
                            bh2.length = (UInt16)(this.StreamReader.ReadUInt16());

                            buffer = new byte[bh2.length];

                            // create a buffer with the bytes from the records and put that array into the 
                            // list 
                            buffer = this.StreamReader.ReadBytes((int)bh2.length);
                            // byteArrayList.Add(buffer);

                            // create for each continue record a new streamreader !! 
                            MemoryStream contbufferstream = new MemoryStream(buffer);
                            VirtualStreamReader contreader = new VirtualStreamReader(contbufferstream);
                            vsrList.AddLast(contreader);


                            // take next Biffrecord ID 
                            bh2.id = (RecordType)this.StreamReader.ReadUInt16();
                        }
                        // set the old position of the stream 
                        this.StreamReader.BaseStream.Position = this.oldOffset;

                        sst = new SST(binreader, bh.id, length, vsrList);
                        this.StreamReader.BaseStream.Position = this.oldOffset + bh.length;
                        this.workBookData.SstData = new SSTData(sst);

                    }

                    else if (bh.id == RecordType.EOF)
                    {
                        // Reads the end of the internal file !!! 
                        this.StreamReader.BaseStream.Seek(0, SeekOrigin.End);
                    }
                    else if (bh.id == RecordType.ExternSheet)
                    {
                        ExternSheet extsheet = new ExternSheet(this.StreamReader, bh.id, bh.length);
                        this.externSheets.Add(extsheet);
                        this.workBookData.addExternSheetData(extsheet);
                    }
                    else if (bh.id == RecordType.SupBook)
                    {
                        SupBook supbook = new SupBook(this.StreamReader, bh.id, bh.length);
                        this.supBooks.Add(supbook);
                        this.workBookData.addSupBookData(supbook);
                    }
                    else if (bh.id == RecordType.XCT)
                    {
                        XCT xct = new XCT(this.StreamReader, bh.id, bh.length);
                        this.XCTList.Add(xct);
                        this.workBookData.addXCT(xct);
                    }
                    else if (bh.id == RecordType.CRN)
                    {
                        CRN crn = new CRN(this.StreamReader, bh.id, bh.length);
                        this.CRNList.Add(crn);
                        this.workBookData.addCRN(crn);
                    }
                    else if (bh.id == RecordType.ExternName)
                    {
                        ExternName externname = new ExternName(this.StreamReader, bh.id, bh.length);
                        this.workBookData.addEXTERNNAME(externname);
                    }
                    else if (bh.id == RecordType.Format)
                    {
                        Format format = new Format(this.StreamReader, bh.id, bh.length);
                        this.workBookData.styleData.addFormatValue(format);
                    }
                    else if (bh.id == RecordType.XF)
                    {
                        XF xf = new XF(this.StreamReader, bh.id, bh.length);
                        this.workBookData.styleData.addXFDataValue(xf);

                    }

                    else if (bh.id == RecordType.Style)
                    {
                        Style style = new Style(this.StreamReader, bh.id, bh.length);
                        this.workBookData.styleData.addStyleValue(style);
                    }
                    else if (bh.id == RecordType.Font)
                    {
                        Font font = new Font(this.StreamReader, bh.id, bh.length);
                        this.workBookData.styleData.addFontData(font);

                    }
                    else if (bh.id == RecordType.NAME)
                    {
                        NAME name = new NAME(this.StreamReader, bh.id, bh.length);
                        this.workBookData.addDefinedName(name);

                    }
                    else if (bh.id == RecordType.BOF)
                    {
                        if (firstBOF)
                        {
                            BOF bof = new BOF(this.StreamReader, bh.id, bh.length);
                        }
                        else
                        {
                            this.StreamReader.ReadBytes(bh.length);
                        }
                    }
                    else if (bh.id == RecordType.FilePass)
                    {
                        throw new ExtractorException(ExtractorException.FILEENCRYPTED);
                    }
                    else if (bh.id == RecordType.Palette)
                    {
                        Palette palette = new Palette(this.StreamReader, bh.id, bh.length);
                        workBookData.styleData.setColorList(palette.rgbColorList); 
                    }
                    else
                    {
                        // this else statement is used to read BiffRecords which aren't implemented 
                        byte[] buffer = new byte[bh.length];
                        buffer = this.StreamReader.ReadBytes(bh.length);
                    }
                }
            }
            catch (Exception ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
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
