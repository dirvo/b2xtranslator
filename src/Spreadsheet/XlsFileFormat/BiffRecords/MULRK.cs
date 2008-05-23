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
    public class MULRK : BiffRecord
    {
        public const RecordNumber ID = RecordNumber.MULRK;

        public UInt16 rw;              // Row 
        public UInt16 colFirst;        // first column 
        public UInt16 colLast;         // Last column  
        List<UInt16> ixfe;             // List Records 
        List<double> rknumber;

        public MULRK(IStreamReader reader, RecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.ixfe = new List<UInt16>();
            this.rknumber = new List<double>(); 

            // count records - 6 standard non variable values !!! 
            int count = (this.Length - 6) / 6 ;
            this.rw = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            for (int i = 0; i < count; i++)
            {
                this.ixfe.Add(reader.ReadUInt16());
                Byte[] buffer = reader.ReadBytes(4);

                rknumber.Add(this.NumFromRK(buffer)); 
            }
            this.colLast = reader.ReadUInt16(); 
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
       }

        public double NumFromRK(Byte[] rk)
        {
            double num=0;
            int high = 1023; 
            UInt32 number; 
            number = System.BitConverter.ToUInt32(rk, 0);
            UInt32 mant = 0;
            // masking the mantisse 
            mant = number & 0x000ffffc;
            // shifting the result by 2  
            mant = mant >> 2; 
            
            UInt32 exp = 0;
            // masking the exponent 
            exp = number & 0x7ff00000;
            // shifting the exponent by 20 
            exp = exp >> 20; 
            // (1 + (Mantisse / 2^18)) * 2 ^ (Exponent - 1023) 
            num = (1 + (mant / System.Math.Pow(2.0, 18.0))) * System.Math.Pow(2, (double)(exp - high)); 
            return num; 
        }
    }
}
