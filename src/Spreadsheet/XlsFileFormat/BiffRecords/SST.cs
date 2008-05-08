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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords
{
    public class SST : BiffRecord
    {
        public const RecordNumber ID = RecordNumber.SST;

        public UInt32 cstTotal;
        public UInt32 cstUnique;

        public List<String> StringList; 

        public SST(IStreamReader reader, RecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.StringList = new List<string>(); 
            // byte[] buffer = new byte[length];
            // buffer = reader.ReadBytes(length);


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

                }
                else if (isRichString)
                {

                }
                else
                {
                    if (isCompressedString)
                    {
                        String buffer = "";
                        for (int j = 0; j < cch; j++)
                        {
                            buffer += (char)reader.ReadByte(); 
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
            back += "Strings: \n" ;
            foreach (String var in this.StringList)
            {
                back += var + "\n"; 
            }

            return back; 
        }
    }
}
