/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords
{
    /// <summary>
    /// This record specifies the text in a text box or a form control.
    /// </summary>
    [BiffRecordAttribute(RecordType.TxO)]
    public class TxO : BiffRecord
    {
        public const RecordType ID = RecordType.TxO;

        public enum HorizontalAlignment : ushort
        {
            Left = 1,
            Centered = 2,
            Right = 3,
            Justify = 4,
            JustifyDistributed = 7
        }

        public enum VerticalAlignment : ushort
        {
            Top = 1,
            Middle = 2,
            Bottom = 3,
            Justify = 4,
            JustifyDistributed = 7
        }

        public enum TextRotation
        {
            Custom,
            Stacked,
            CounterClockwise,
            Clockwise
        }

        /// <summary>
        /// Specifies the horizontal alignment.
        /// </summary>
        public HorizontalAlignment hAlignment;

        /// <summary>
        /// Specifies the vertical alignment.
        /// </summary>
        public VerticalAlignment vAlignment;

        /// <summary>
        /// Specifies the orientation of the text within the object boundary.
        /// </summary>
        public TextRotation rot;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in the text string 
        /// contained in the Continue records immediately following this record. <br/>
        /// MUST be less than or equal to 255.
        /// </summary>
        public UInt16 cchText;

        /// <summary>
        /// An unsigned integer that specifies the number of bytes of formatting run information in the 
        /// TxORuns structure contained in the Continue records following this record.<br/>
        /// If cchText is 0, this value MUST be 0.<br/>
        /// Otherwise the value MUST be greater than or equal to 16 and MUST be a multiple of 8.
        /// </summary>
        public UInt16 cbRuns;

        /// <summary>
        /// A FontIndex that specifies the font when cchText is 0.<br/>
        /// </summary>
        public UInt16 ifntEmpty;



        public TxO(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            // initialize class members from stream
            UInt16 flags = reader.ReadUInt16();
            this.hAlignment = (HorizontalAlignment)Utils.BitmaskToInt(flags, 0xE);
            this.vAlignment = (VerticalAlignment)Utils.BitmaskToInt(flags, 0x70);
            this.rot = (TextRotation)reader.ReadUInt16();
            reader.ReadBytes(6); // reserved
            // TODO: Check for missing field controlInfo in MS-XLS
            this.cchText = reader.ReadUInt16();
            this.cbRuns = reader.ReadUInt16();
            this.ifntEmpty = reader.ReadUInt16();
            reader.ReadBytes(2); // reserved

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
