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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This class extracts the SST-Record Data from the specific biffrecord 
    /// </summary>
    [BiffRecordAttribute(RecordType.SST)] 
    public class SST : BiffRecord
    {
        public LinkedList<VirtualStreamReader> contStreamlist; 
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
        public SST(IStreamReader binreader, RecordType id, UInt16 length, LinkedList<VirtualStreamReader> contstreamlist)
            : base(binreader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.contStreamlist = contstreamlist; 
            this.StringList = new List<string>();
            this.FormatList = new List<StringFormatAssignment>();
            byte[] buffer = new byte[length];
            int counti = 0;
            this.cstTotal = (UInt32)this.Reader.ReadUInt32();
            this.cstUnique = this.Reader.ReadUInt32();
            
            try
            {
                // run over the different strings 
                // there are x strings where x = cstUnique    
                for (int i = 0; i < this.cstUnique; i++)
                {
                    counti++;
                    if (this.Reader.BaseStream.Position == this.Reader.BaseStream.Length )
                    {
                        this.switchStream(); 
                    }
                    // first get the char count of this string 
                    UInt16 cch = this.Reader.ReadUInt16();
                    // get the grbit mask 
                    Byte grbit = this.Reader.ReadByte();
                    bool isCompressedString = false;
                    bool isExtString = false;
                    bool isRichString = false;

                    int cRun = 0;
                    int cbExtRst = 0; 

                    // demask the grbit 
                    isCompressedString = !Utils.BitmaskToBool((int)grbit, 0x0001);
                    isExtString = Utils.BitmaskToBool((int)grbit, 0x0004);
                    isRichString = Utils.BitmaskToBool((int)grbit, 0x0008);



                    if (isRichString)
                    {
                        cRun = this.Reader.ReadUInt16(); 
                    }

                    if (isExtString)
                    {
                        cbExtRst = this.Reader.ReadInt32(); 
                    }

                    // switch stream  
                    if (this.Reader.BaseStream.Position == this.Reader.BaseStream.Length)
                    {
                        this.switchStream();
                    }


                    // read characters from the string
                    int charcount = 0;
                    if (isCompressedString)
                        charcount = 1;
                    else
                        charcount = 2;
                    String stringbuffer = "";
                    // read chars !!! 
                    while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + cch * charcount)
                    {
                        ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                        cch -= (ushort)(currentLength / charcount);
                        for (int j = 0; j < currentLength / charcount; j++)
                        {
                            if (isCompressedString)
                            {
                                stringbuffer += (char)this.Reader.ReadByte();
                            }
                            else
                            {
                                stringbuffer += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                            }
                        }
                        // switch to next stream !! 
                        this.switchStream();
                        // read compressed/uncompressed byte value 
                        byte grbit2 = this.Reader.ReadByte();
                        if (grbit2 > 0)
                            isCompressedString = false;
                        else
                            isCompressedString = true; 
                    }

                    for (int j = 0; j < cch; j++)
                    {
                        if (isCompressedString)
                        {
                            stringbuffer += (char)this.Reader.ReadByte();
                        }
                        else
                        {
                            stringbuffer += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                        }
                    }
                    this.StringList.Add(stringbuffer);

                    // read formatting runs!! 
                    if (isRichString)
                    {
                        int countFormatingRuns = cRun; 
                        while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + countFormatingRuns * 4)
                        {
                            ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                            countFormatingRuns -= (ushort)(currentLength / 4);
                            // get formating data 
                            for (int j = 0; j < currentLength / 4; j++)
                            {
                                StringFormatAssignment format = new StringFormatAssignment();
                                format.StringNumber = counti;
                                format.CharNumber = this.Reader.ReadUInt16();



                                format.FontRecord = this.Reader.ReadUInt16();
                                if (format.CharNumber < stringbuffer.Length)
                                {
                                    this.FormatList.Add(format);
                                }
                            }
                            // switch to next stream !! 
                            this.switchStream();
                            // read compressed/uncompressed byte value 
                            // this.Reader.ReadByte();
                        }
                        // get formating data 
                        for (int j = 0; j < countFormatingRuns; j++)
                        {
                            StringFormatAssignment format = new StringFormatAssignment();
                            format.StringNumber = counti;
                            format.CharNumber = this.Reader.ReadUInt16();
                            format.FontRecord = this.Reader.ReadUInt16();

                            if (format.FontRecord > 4)
                                format.FontRecord--; 

                            /// ToDo: Check why some charNumbers are greater then string length 
                            if (format.CharNumber < stringbuffer.Length)
                            {
                                this.FormatList.Add(format);
                            }
                        }
                    }

                    if (isExtString)
                    {
                        int cchExtRst = cbExtRst;
                        byte[] ExtRst;
                        while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + cchExtRst)
                        {
                            ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                            cchExtRst -= (currentLength);
                            ExtRst = this.Reader.ReadBytes(currentLength);
                            // switch to next stream !! 
                            this.switchStream();
                            // read compressed/uncompressed byte value 
                            // this.Reader.ReadByte();
                        }

                        ExtRst = this.Reader.ReadBytes(cchExtRst);
                    }

                    
                }
            }
            catch (Exception ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
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

        /// <summary>
        /// This method is used to switch a stream, if one stream is at the end 
        /// the stream has to be changed 
        /// </summary>
        public void switchStream()
        {
            if (this.contStreamlist.Count > 0)
            {
                this.Reader = (VirtualStreamReader)this.contStreamlist.First.Value;
                this.contStreamlist.RemoveFirst();
            }
        }
    }
}
