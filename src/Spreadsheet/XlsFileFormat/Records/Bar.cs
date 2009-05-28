﻿/*
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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords.Graph
{
    /// <summary>
    /// This record specifies that the chart group is a bar chart group or 
    /// a column chart group, and specifies the chart group attributes.
    /// </summary>
    [BiffRecordAttribute(RecordType.Bar)]
    public class Bar : BiffRecord
    {
        public const RecordType ID = RecordType.Bar;

        /// <summary>
        /// A signed integer that specifies the overlap between data points in the same category 
        /// as a percentage of the data point width. 
        /// 
        /// MUST be greater than or equal to -100 and less than or equal to 100. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value         Meaning
        ///     -100 to -1    Size of the separation between data points
        ///     0             No overlap
        ///     1 to 100      Size of the overlap between data points
        /// </summary>
        public Int16 pcOverlap;
        
        /// <summary>
        /// An unsigned integer that specifies the width of the gap between the categories 
        /// and the left and right edges of the plot area as a percentage of the data point width divided by 2. 
        /// 
        /// It also specifies the width of the gap between adjacent categories 
        /// as a percentage of the data point width. 
        /// 
        /// MUST be less than or equal to 500.
        /// </summary>
        public UInt16 pcGap;

        /// <summary>
        /// A bit that specifies whether the data points and value axis are horizontal 
        /// (for a bar chart group) or vertical (for a column chart group). 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value   Meaning
        ///     0       Data points and value axis are vertical.
        ///     1       Data points and value axis are horizontal.
        /// </summary>
        public bool fTranspose;

        /// <summary>
        /// A bit that specifies whether the data points in the chart group 
        /// that share the same category are stacked one on top of the next.
        /// </summary>
        public bool fStacked;

        /// <summary>
        /// A bit that specifies whether the data points in the chart group are displayed 
        /// as a percentage of the sum of all data points in the chart group 
        /// that share the same category. 
        /// 
        /// MUST be 0 if fStacked is 0.
        /// </summary>
        public bool f100;

        /// <summary>
        /// A bit that specifies whether one or more data points in the chart group has shadows.
        /// </summary>
        public bool fHasShadow;

        public Bar(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.pcOverlap = reader.ReadInt16();
            this.pcGap = reader.ReadUInt16();

            UInt16 flags = reader.ReadUInt16();
            this.fTranspose = Utils.BitmaskToBool(flags, 0x1);
            this.fStacked = Utils.BitmaskToBool(flags, 0x2);
            this.f100 = Utils.BitmaskToBool(flags, 0x4);
            this.fHasShadow = Utils.BitmaskToBool(flags, 0x8);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
