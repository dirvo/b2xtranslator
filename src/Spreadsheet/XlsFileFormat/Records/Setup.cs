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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.Setup)] 
    public class Setup : BiffRecord
    {
        public const RecordType ID = RecordType.Setup;

        public UInt16 iPaperSize;
        public UInt16 iScale;
        public UInt16 iPageStart;
        public UInt16 iFitWidth;
        public UInt16 iFitHeight;
        public UInt16 grbit;
        public UInt16 iRes;
        public UInt16 iVRes;
        public UInt16 iCopies;

        public double numHdr;
        public double numFtr;

        public bool fLeftToRight;
        public bool fPortrait;
        public bool fNoPls;
        public bool fNoColor;

        public bool fDraft;
        public bool fNotes;
        public bool fNoOrient;
        public bool fUsePage;

        public bool fEndNotes;
        public int iErrors; 

        public Setup(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);


            this.iPaperSize = reader.ReadUInt16();
            this.iScale = reader.ReadUInt16();
            this.iPageStart = reader.ReadUInt16();
            this.iFitWidth = reader.ReadUInt16();
            this.iFitHeight = reader.ReadUInt16();
            this.grbit = reader.ReadUInt16();
            this.iRes = reader.ReadUInt16();
            this.iVRes = reader.ReadUInt16();

            this.numHdr = reader.ReadDouble();
            this.numFtr = reader.ReadDouble(); 

            this.iCopies = reader.ReadUInt16(); 

            // set flags 
            this.fLeftToRight = Utils.BitmaskToBool(this.grbit, 0x01);
            this.fPortrait = Utils.BitmaskToBool(this.grbit, 0x02);
            this.fNoPls = Utils.BitmaskToBool(this.grbit, 0x04);
            this.fNoColor = Utils.BitmaskToBool(this.grbit, 0x08);

            this.fDraft = Utils.BitmaskToBool(this.grbit, 0x010);
            this.fNotes = Utils.BitmaskToBool(this.grbit, 0x020);
            this.fNoOrient = Utils.BitmaskToBool(this.grbit, 0x040);
            this.fUsePage = Utils.BitmaskToBool(this.grbit, 0x080);

            this.fEndNotes = Utils.BitmaskToBool(this.grbit, 0x080);
            this.iErrors = (this.grbit & 0x0C00) << 0x0A;
        }
    }
}
