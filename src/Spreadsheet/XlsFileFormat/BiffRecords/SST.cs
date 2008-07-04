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
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords
{
    /// <summary>
    /// This class extracts the SST-Record Data from the specific biffrecord 
    /// </summary>
    public class SST : BiffRecord
    {
        public LinkedList<VirtualStreamReader> contStreamlist; 
        /// <summary>
        /// the own record data id 
        /// </summary>
        public const RecordNumber ID = RecordNumber.SST;

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
        public SST(IStreamReader binreader, RecordNumber id, UInt32 length, LinkedList<VirtualStreamReader> contstreamlist)
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
                        Console.WriteLine("Test");
                        this.switchStream(); 
                    }
                    // first get the char count of this string 
                    UInt16 cch = this.Reader.ReadUInt16();
                    // get the grbit mask 
                    Byte grbit = this.Reader.ReadByte();
                    bool isCompressedString = false;
                    bool isExtString = false;
                    bool isRichString = false;

                    // demask the grbit 
                    isCompressedString = !Utils.BitmaskToBool((int)grbit, 0x0001);
                    isExtString = Utils.BitmaskToBool((int)grbit, 0x0004);
                    isRichString = Utils.BitmaskToBool((int)grbit, 0x0008);

                    if (isExtString)
                    {
                        // Two versions, first is extended string and no rich string 
                        // second is extended and rich string 
                        // first 
                        if (!isRichString)
                        {
                            Int32 cchExtRst = this.Reader.ReadInt32();
                            String stringbuffer = "";
                            while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + cch * 2)
                            {
                                ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                                cch -= (ushort)(currentLength / 2);
                                for (int j = 0; j < currentLength / 2; j++)
                                {
                                    stringbuffer += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                                }
                                // switch to next stream !! 
                                this.switchStream();
                                // read compressed/uncompressed byte value 
                                this.Reader.ReadByte(); 
                            }
                            for (int j = 0; j < cch; j++)
                            {
                                stringbuffer += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                            }
                            this.StringList.Add(stringbuffer);
                            byte[] ExtRst; 
                            // read undocumented data structure ExtRst 
                            while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + cchExtRst)
                            {
                                ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                                cchExtRst -= (currentLength);
                                ExtRst = this.Reader.ReadBytes(currentLength); 
                                // switch to next stream !! 
                                this.switchStream();
                                // read compressed/uncompressed byte value 
                                this.Reader.ReadByte();
                            }
                            ExtRst = this.Reader.ReadBytes(cchExtRst);
                        } //second 
                        else
                        {
                            String stringbuffer = "";
                            UInt16 countFormatingRuns = this.Reader.ReadUInt16();
                            Int32 cchExtRst = this.Reader.ReadInt32();
                            // read chars !!! 
                            while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + cch * 2)
                            {
                                ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                                cch -= (ushort)(currentLength / 2);
                                for (int j = 0; j < currentLength / 2; j++)
                                {
                                    stringbuffer += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                                }
                                // switch to next stream !! 
                                this.switchStream();
                                // read compressed/uncompressed byte value 
                                this.Reader.ReadByte();
                            }
                           
                            for (int j = 0; j < cch; j++)
                            {
                                stringbuffer += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                            }
                            this.StringList.Add(stringbuffer);
                            byte[] rgSTRUN;
                            byte[] ExtRst; 
                            // read rgSTRUN 
                            while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + countFormatingRuns * 4)
                            {
                                ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                                countFormatingRuns -= (ushort)(currentLength / 4);
                                rgSTRUN = this.Reader.ReadBytes(currentLength);
                                // switch to next stream !! 
                                this.switchStream();
                                // read compressed/uncompressed byte value 
                                this.Reader.ReadByte();
                            }
                            // read formating run structures 
                            rgSTRUN = this.Reader.ReadBytes(countFormatingRuns * 4);

                            while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + cchExtRst)
                            {
                                ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                                cchExtRst -= (currentLength);
                                ExtRst = this.Reader.ReadBytes(currentLength);
                                // switch to next stream !! 
                                this.switchStream();
                                // read compressed/uncompressed byte value 
                                this.Reader.ReadByte();
                            }

                            ExtRst = this.Reader.ReadBytes(cchExtRst);

                        }
                    }
                    // Rich strings are formated string values, there are some more information 
                    else if (isRichString && !isExtString)
                    {
                        // get number of formating runs !! 
                        UInt16 countFormatingRuns = this.Reader.ReadUInt16();
                        // String buffer = Encoding.Unicode.GetString(reader.ReadBytes(cch));

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
                            // 
                        }
                        this.StringList.Add(stringbuffer);

                        while (this.Reader.BaseStream.Length < this.Reader.BaseStream.Position + countFormatingRuns * 4)
                        {
                            ushort currentLength = (ushort)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position);
                            countFormatingRuns -= (ushort)(currentLength / 4);
                            // get formating data 
                            for (int j = 0; j < currentLength/4; j++)
                            {
                                StringFormatAssignment format = new StringFormatAssignment();
                                format.StringNumber = i;
                                format.CharNumber = this.Reader.ReadUInt16();
                                format.FontRecord = this.Reader.ReadUInt16();
                                this.FormatList.Add(format);
                            }
                            // switch to next stream !! 
                            this.switchStream();
                            // read compressed/uncompressed byte value 
                            this.Reader.ReadByte();
                        }
                        // get formating data 
                        for (int j = 0; j < countFormatingRuns; j++)
                        {
                            StringFormatAssignment format = new StringFormatAssignment();
                            format.StringNumber = i;
                            format.CharNumber = this.Reader.ReadUInt16();
                            format.FontRecord = this.Reader.ReadUInt16();
                            this.FormatList.Add(format);
                        }
                    }
                    else
                    {
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
                            // 
                        }
                        this.StringList.Add(stringbuffer);                       
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
            this.Reader = (VirtualStreamReader)this.contStreamlist.First.Value;
            this.contStreamlist.RemoveFirst(); 
        }
    }
}
