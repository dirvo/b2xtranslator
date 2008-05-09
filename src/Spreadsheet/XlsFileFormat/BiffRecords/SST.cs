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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords
{
    /// <summary>
    /// This class extracts the SST-Record Data from the specific biffrecord 
    /// </summary>
    public class SST : BiffRecord
    {
        /// <summary>
        /// a simple struct to hold the format data from strings 
        /// </summary>
        public struct StringFormatAssignment
        {
            public int StringNumber;
            public UInt16 CharNumber;
            public UInt16 FontRecord;
        }

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
        public SST(IStreamReader reader, RecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.StringList = new List<string>();
            this.FormatList = new List<StringFormatAssignment>();


            this.cstTotal = (UInt32)reader.ReadUInt32();
            this.cstUnique = reader.ReadUInt32();

            // run over the different strings 
            // there are x strings where x = cstUnique 
            for (int i = 0; i < this.cstUnique; i++)
            {
                // first get the char count of this string 
                UInt16 cch = reader.ReadUInt16();
                // get the grbit mask 
                Byte grbit = reader.ReadByte();

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
                        Int32 cchExtRst = reader.ReadInt32();
                        String buffer = "";
                        for (int j = 0; j < cch; j++)
                        {
                            buffer += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
                        }
                        this.StringList.Add(buffer);

                        // read undocumented data structure ExtRst 
                        byte[] ExtRst = reader.ReadBytes(cchExtRst);
                    } //second 
                    else
                    {
                        UInt16 countFormatingRuns = reader.ReadUInt16();
                        Int32 cchExtRst = reader.ReadInt32();
                        String buffer = "";
                        for (int j = 0; j < cch; j++)
                        {
                            buffer += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
                        }
                        this.StringList.Add(buffer);
                        // read formating run structures 
                        byte[] rgSTRUN = reader.ReadBytes(countFormatingRuns * 4);
                        byte[] ExtRst = reader.ReadBytes(cchExtRst);

                    }
                }
                // Rich strings are formated string values, there are some more information 
                else if (isRichString && !isExtString)
                {
                    // get number of formating runs !! 
                    UInt16 countFormatingRuns = reader.ReadUInt16();
                    String buffer = Encoding.Unicode.GetString(reader.ReadBytes(cch));
                    
                    for (int j = 0; j < cch; j++)
                    {
                        buffer += (char)reader.ReadByte();
                    }
                    this.StringList.Add(buffer);
                    // get formating data 
                    for (int j = 0; j < countFormatingRuns; j++)
                    {
                        StringFormatAssignment format;
                        format.StringNumber = i;
                        format.CharNumber = reader.ReadUInt16();
                        format.FontRecord = reader.ReadUInt16();
                        this.FormatList.Add(format);
                    }
                }
                else
                {
                    // compressed strings are strings which use only one byte per character 
                    if (isCompressedString)
                    {
                        String buffer = "";
                        for (int j = 0; j < cch; j++)
                        {
                            buffer += (char)reader.ReadByte();
                        }
                        this.StringList.Add(buffer);
                    }
                    // not compressed strings are two bytes long 
                    else
                    {
                        String buffer = "";
                        for (int j = 0; j < cch; j++)
                        {
                            buffer += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
                        }
                        this.StringList.Add(buffer);
                    }
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

   
    }
}
