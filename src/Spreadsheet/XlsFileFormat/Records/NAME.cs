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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.NAME)] 
    public class NAME : BiffRecord
    {
        public const RecordType ID = RecordType.NAME;

        public byte chKey;
        public byte cch;
        public UInt16 cce;
        public UInt16 itab;
        public string Name; 
        public Stack<AbstractPtg> ptgStack; 

        public NAME(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            int options = reader.ReadUInt16();
            this.chKey = reader.ReadByte();
            this.cch = reader.ReadByte();
            this.cce = reader.ReadUInt16(); 
            //read reserved bytes 
            reader.ReadBytes(2);
            this.itab = reader.ReadUInt16(); 
            // read 4 reserved bytes 
            reader.ReadBytes(4);

            byte grbit = reader.ReadByte();
            this.Name = ExcelHelperClass.getStringFromBiffRecord(reader, (int)this.cch, (int)grbit);




            long oldStreamPosition = this.Reader.BaseStream.Position;
            try
            {
                this.ptgStack = ExcelHelperClass.getFormulaStack(this.Reader, this.cce);
            }
            catch (Exception ex)
            {
                this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                this.Reader.BaseStream.Seek(this.cce, System.IO.SeekOrigin.Current);
                TraceLogger.Debug("Formula parse error in intern name");
                TraceLogger.Error(ex.StackTrace);
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
