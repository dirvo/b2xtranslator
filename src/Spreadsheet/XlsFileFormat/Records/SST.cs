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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;
using System.IO;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This class extracts the SST-Record Data from the specific biffrecord 
    /// </summary>
    [BiffRecordAttribute(RecordType.SST)] 
    public class SST : BiffRecord, IVisitable
    {
        /// <summary>
        /// the own record data id 
        /// </summary>
        public const RecordType ID = RecordType.SST;

        /// <summary>
        /// Total and unique number of strings in this SST-Biffrecord 
        /// </summary>
        public UInt32 cstTotal;
        public UInt32 cstUnique;

        public List<String> StringList;
        public List<StringFormatAssignment> FormatList;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Reader to parse the document </param>
        /// <param name="id">BiffRecord ID</param>
        /// <param name="length">The lengt of the biffrecord </param>
        public SST(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            this.StringList = new List<string>();
            this.FormatList = new List<StringFormatAssignment>();

            // copy data to a memory stream to cope with Continue records
            using (MemoryStream ms = new MemoryStream())
            {
                List<long> recordBorders = new List<long>();
                
                byte[] buffer = reader.ReadBytes(length);
                ms.Write(buffer, 0, length);

                while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
                {
                    BiffHeader bh;
                    bh.id = (RecordType)reader.ReadUInt16();
                    bh.length = reader.ReadUInt16();

                    recordBorders.Add(ms.Position);
                    buffer = reader.ReadBytes(bh.length);
                    ms.Write(buffer, 0, bh.length);
                }

                ms.Position = 0;

                VirtualStreamReader sr = new VirtualStreamReader(ms);
                
                try
                {
                    this.cstTotal = sr.ReadUInt32();
                    this.cstUnique = sr.ReadUInt32();
            
                    // run over the different strings 
                    // there are x strings where x = cstUnique    
                    for (int i = 0; i < this.cstUnique; i++)
                    {
                        XLUnicodeRichExtendedString str = new XLUnicodeRichExtendedString(sr, recordBorders);

                        //// first get the char count of this string 
                        //UInt16 cch = sr.ReadUInt16();
                        //// get the grbit mask 
                        //Byte grbit = sr.ReadByte();
                        //bool isCompressedString = !Utils.BitmaskToBool((int)grbit, 0x0001);
                        //bool isExtString = Utils.BitmaskToBool((int)grbit, 0x0004);
                        //bool isRichString = Utils.BitmaskToBool((int)grbit, 0x0008);

                        //int cRun = 0;
                        //int cbExtRst = 0;

                        //if (isRichString)
                        //{
                        //    cRun = sr.ReadUInt16();
                        //}

                        //if (isExtString)
                        //{
                        //    cbExtRst = sr.ReadInt32();
                        //}

                        //// read characters from the string
                        //int charcount = 0;
                        //if (isCompressedString)
                        //    charcount = 1;
                        //else
                        //    charcount = 2;

                        //String stringbuffer = "";

                        //// read chars !!! 
                        //while (sr.BaseStream.Length < sr.BaseStream.Position + cch * charcount)
                        //{
                        //    ushort currentLength = (ushort)(sr.BaseStream.Length - sr.BaseStream.Position);
                        //    cch -= (ushort)(currentLength / charcount);
                        //    for (int j = 0; j < currentLength / charcount; j++)
                        //    {
                        //        if (isCompressedString)
                        //        {
                        //            stringbuffer += (char)sr.ReadByte();
                        //        }
                        //        else
                        //        {
                        //            stringbuffer += System.BitConverter.ToChar(sr.ReadBytes(2), 0);
                        //        }
                        //    }
                            
                        //    // read compressed/uncompressed byte value 
                        //    byte grbit2 = sr.ReadByte();
                        //    if (grbit2 > 0)
                        //        isCompressedString = false;
                        //    else
                        //        isCompressedString = true;
                        //}

                        //for (int j = 0; j < cch; j++)
                        //{
                        //    if (isCompressedString)
                        //    {
                        //        stringbuffer += (char)sr.ReadByte();
                        //    }
                        //    else
                        //    {
                        //        stringbuffer += System.BitConverter.ToChar(sr.ReadBytes(2), 0);
                        //    }
                        //}
                        this.StringList.Add(str.Value);

                        if (str.fRichSt)
                        {
                            foreach (FormatRun formatRun in str.rgRun)
                            {
                                StringFormatAssignment format = new StringFormatAssignment();
                                format.StringNumber = i + 1;

                                format.CharNumber = formatRun.ich;
                                format.FontRecord = formatRun.ifnt;

                                // NOTE: If this value is less than 4, then it specifies a *zero-based* index of a Font record 
                                //  in the collection of Font records in the globals substream. If this value is greater than 4, 
                                //  then it specifies a *one-based* index of a Font record in the collection of Font records in 
                                //  the globals substream. 
                                //
                                if (format.FontRecord > 4)
                                {
                                    format.FontRecord--;
                                }

                                if (format.CharNumber < str.Value.Length)
                                {
                                    this.FormatList.Add(format);
                                }
                            }
                        }

                        //// read formatting runs!! 
                        //if (isRichString)
                        //{
                        //    int countFormatingRuns = cRun;
                        //    while (sr.BaseStream.Length < sr.BaseStream.Position + countFormatingRuns * 4)
                        //    {
                        //        ushort currentLength = (ushort)(sr.BaseStream.Length - sr.BaseStream.Position);
                        //        countFormatingRuns -= (ushort)(currentLength / 4);
                        //        // get formating data 
                        //        for (int j = 0; j < currentLength / 4; j++)
                        //        {
                        //            StringFormatAssignment format = new StringFormatAssignment();
                        //            format.StringNumber = i + 1;

                        //            format.CharNumber = sr.ReadUInt16();
                        //            format.FontRecord = sr.ReadUInt16();
                                    
                        //            if (format.CharNumber < stringbuffer.Length)
                        //            {
                        //                this.FormatList.Add(format);
                        //            }
                        //        }
                        //    }
                        //    // get formating data 
                        //    for (int j = 0; j < countFormatingRuns; j++)
                        //    {
                        //        StringFormatAssignment format = new StringFormatAssignment();
                        //        format.StringNumber = i + 1;
                        //        format.CharNumber = sr.ReadUInt16();
                        //        format.FontRecord = sr.ReadUInt16();

                        //        if (format.FontRecord > 4)
                        //        {
                        //            format.FontRecord--;
                        //        }

                        //        /// ToDo: Check why some charNumbers are greater then string length 
                        //        if (format.CharNumber < stringbuffer.Length)
                        //        {
                        //            this.FormatList.Add(format);
                        //        }
                        //    }
                        //}

                        //if (isExtString)
                        //{
                        //    int cchExtRst = cbExtRst;
                        //    byte[] ExtRst;
                        //    while (sr.BaseStream.Length < sr.BaseStream.Position + cchExtRst)
                        //    {
                        //        ushort currentLength = (ushort)(sr.BaseStream.Length - sr.BaseStream.Position);
                        //        cchExtRst -= (currentLength);
                        //        ExtRst = sr.ReadBytes(currentLength);
                        //    }

                        //    ExtRst = sr.ReadBytes(cchExtRst);
                        //}
                    }
                }
                catch (Exception ex)
                {
                    TraceLogger.Error(ex.Message);
                    TraceLogger.Debug(ex.ToString());
                }
            }
            // assert that the correct number of bytes has been read from the stream
            // Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        /// <summary>
        /// The ToString Method
        /// </summary>
        /// <returns></returns>
        public override String ToString()
        {
            String back = "";
            back += "Number Strings Total: " + this.cstTotal + "\n";
            back += "Number Unique Strings: " + this.cstUnique + "\n";
            back += "Strings: \n";
            foreach (String var in this.StringList)
            {
                back += var + "\n";
            }

            return back;
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SST>)mapping).Apply(this);
        }

        #endregion
    }
}
