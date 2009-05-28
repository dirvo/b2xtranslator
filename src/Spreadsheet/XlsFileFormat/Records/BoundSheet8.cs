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
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords
{
    /// <summary>
    /// BOUNDSHEET: Sheet Information (85h)
    /// 
    /// This record stores the sheet name, sheet type, and stream position.
    /// </summary>
    [BiffRecordAttribute(RecordType.BoundSheet8)] 
    public class BoundSheet8 : BiffRecord
    {
        public const RecordType ID = RecordType.BoundSheet8;

        /// <summary>
        /// Some enum definitions 
        /// </summary>
        public enum hiddenFlags:int {visible=0, hidden=1, veryhidden=2 };
        public enum sheetTypes:int { worksheet=0, macrosheet=1, chart=2, visualbasic=6 }; 


        /// <summary>
        /// Stream position of the start of the BOF record for the sheet
        /// </summary>
        public UInt32 lbPlyPos;

        /// <summary>
        /// Option flags
        /// </summary>
        private UInt16 grbit;

        /// <summary>
        /// Length of the sheet name (in characters)
        /// </summary>
        public byte cch;

        /// <summary>
        /// Sheet name (grbit/rgb fields of Unicode String)
        /// </summary>
        public byte[] rgch;
        // TODO: check for correct interpretation of Unicode strings

        /// <summary>
        /// The hidden status of the workbook 
        /// </summary>
        public int hiddenState;

        /// <summary>
        /// The sheet type value
        /// </summary>
        public sheetTypes sheetType; 

        /// <summary>
        /// extracts the boundsheetdata from the biffrecord  
        /// </summary>
        /// <param name="reader">IStreamReader </param>
        /// <param name="id">Type of the record </param>
        /// <param name="length">Length of the record</param>
        public BoundSheet8(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            this.lbPlyPos = this.Reader.ReadUInt32(); 
            this.grbit = this.Reader.ReadUInt16();
           
            this.cch = this.Reader.ReadByte(); 

            this.rgch = new byte[this.cch];
            
            byte fHighByte = this.Reader.ReadByte();
            int isCompressed = fHighByte & 0x0001;
            for (int i = 0; i < this.cch; i++)
            {
                this.rgch[i] = this.Reader.ReadByte(); 
            }
            
            
            // Setting the hidden state value 
            // Bitmask is 0003h -> first two bits 
            this.hiddenState = Utils.BitmaskToInt(this.grbit, 0x0003); 

            // Setting the sheet type value 
            this.sheetType = (sheetTypes)Utils.BitmaskToInt(this.grbit, 0xFF00);
            
            // assert that the correct number of bytes has been read from the stream
            // Debug.trace(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        /// <summary>
        /// Simple ToString Method 
        /// </summary>
        /// <returns>String from the object</returns>
        public override String ToString()
        {
            String returnvalue = "BOUNDSHEET - RECORD: \n";
            returnvalue += "-- Name: " + this.getBoundsheetName() + "\n";
            returnvalue += "-- Offset: " + this.lbPlyPos + "\n";
            returnvalue += "-- HiddenState: " + (hiddenFlags)this.hiddenState + "\n";
            returnvalue += "-- Sheettype: " + (sheetTypes)this.sheetType + "\n"; 
            return returnvalue; 
        }


        /// <summary>
        /// Helper Method to get the boundsheetname 
        /// </summary>
        /// <returns>Boundsheetname </returns>
        public String getBoundsheetName()
        {
            String returnvalue = ""; 
            for (int i = 0; i < this.rgch.Length; i++)
            {
                returnvalue += (char)this.rgch[i];
            }
            return returnvalue; 
        }

    }
}
