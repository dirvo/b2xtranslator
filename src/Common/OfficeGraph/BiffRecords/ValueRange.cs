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

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the properties of a value axis.
    /// </summary>
    public class ValueRange : OfficeGraphBiffRecord
    {
        public const RecordNumber ID = RecordNumber.ValueRange;

        /// <summary>
        /// An Xnum that specifies the minimum value of the value axis. 
        /// 
        /// MUST be less than numMax. If the value of fAutoMin is 1, this field MUST be ignored.
        /// </summary>
        public double numMin;

        /// <summary>
        /// An Xnum that specifies the maximum value of the value axis. 
        /// 
        /// MUST be greater than numMin. If the value of fAutoMax is 1, this field MUST be ignored.
        /// </summary>
        public double numMax;

        /// <summary>
        /// An Xnum that specifies the interval at which major tick marks and major gridlines are displayed. 
        /// 
        /// MUST be greater than or equal to numMinor. If the value of fAutoMajor is 1, this field MUST be ignored.
        /// </summary>
        public double numMajor;

        /// <summary>
        /// An Xnum that specifies the interval at which minor tick marks and minor gridlines are displayed. 
        /// 
        /// MUST be greater than or equal to zero. If the value of fAutoMinor is 1, this field MUST be ignored.
        /// </summary>
        public double numMinor;

        /// <summary>
        /// An Xnum that specifies at which value the other axes in the axis group cross this value axis. 
        /// 
        /// If the value of fAutoCross is 1, this field MUST be ignored.
        /// </summary>
        public double numCross;

        /// <summary>
        /// A bit that specifies whether numMin is calculated automatically. 
        /// 
        /// MUST be one of the following: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by numMin is used as the minimum value of the value axis.
        ///     1           numMin is calculated such that the data point with the minimum value can be displayed in the plot area.
        /// </summary>
        public bool fAutoMin;

        /// <summary>
        /// A bit that specifies whether numMax is calculated automatically. 
        /// 
        /// MUST be one of the following: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by numMax is used as the maximum value of the value axis.
        ///     1           numMax is calculated such that the data point with the maximum value can be displayed in the plot area.
        /// </summary>
        public bool fAutoMax;

        /// <summary>
        /// A bit that specifies whether numMajor is calculated automatically. 
        /// 
        /// MUST be one of the following: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by numMajor is used as the interval at which major tick marks and major gridlines are displayed.
        ///     1           numMajor is calculated automatically.
        /// </summary>
        public bool fAutoMajor;

        /// <summary>
        /// A bit that specifies whether numMinor is calculated automatically. 
        /// 
        /// MUST be one of the following: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by numMinor is used as the interval at which minor tick marks and minor gridlines are displayed.
        ///     1           numMinor is calculated automatically.
        /// </summary>
        public bool fAutoMinor;

        /// <summary>
        /// A bit that specifies whether numCross is calculated automatically. 
        /// 
        /// MUST be one of the following: 
        ///     
        ///     Value       Meaning
        ///     0           The value specified by numCross is used as the point at which the 
        ///                 other axes in the axis group cross this value axis.
        ///     1           numCross is calculated so that the crossing point is displayed in the plot area.
        /// </summary>
        public bool fAutoCross;

        /// <summary>
        /// A bit that specifies whether the value axis has a logarithmic scale. 
        /// 
        /// MUST be one of the following: 
        /// 
        ///     Value       Meaning
        ///     0           The scale of the value axis is linear.
        ///     1           The scale of the value axis is logarithmic in base 10.
        /// </summary>
        public bool fLog;

        /// <summary>
        /// A bit that specifies whether the values on the value axis are displayed in reverse order. 
        /// 
        /// MUST be one of the following: 
        /// 
        ///     Value       Meaning
        ///     0           Values are displayed from smallest-to-largest from left-to-right or 
        ///                 bottom-to-top, respectively, depending on the orientation of the axis.
        ///     1           The values are displayed in reverse order, meaning largest-to-smallest 
        ///                 from left-to-right or bottom-to-top, respectively.
        /// </summary>
        public bool fReversed;

        /// <summary>
        /// A bit that specifies whether the other axes in the axis group cross this value axis at the maximum value. 
        /// 
        /// MUST be one of the following:
        /// 
        ///     Value       Meaning
        ///     0           The other axes in the axis group cross this value axis at the value specified by numCross.
        ///     1           The other axes in the axis group cross the value axis at the maximum value. 
        ///                 If fMaxCross is 1, then both fAutoCross and numCross MUST be ignored.
        /// </summary>
        public bool fMaxCross;


        public ValueRange(IStreamReader reader, RecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.numMin = reader.ReadDouble();
            this.numMax = reader.ReadDouble();
            this.numMajor = reader.ReadDouble();
            this.numMinor = reader.ReadDouble();
            this.numCross = reader.ReadDouble();

            UInt16 flags = reader.ReadUInt16();
            this.fAutoMin = Utils.BitmaskToBool(flags, 0x0001);
            this.fAutoMax = Utils.BitmaskToBool(flags, 0x0002);
            this.fAutoMajor = Utils.BitmaskToBool(flags, 0x0004);
            this.fAutoMinor = Utils.BitmaskToBool(flags, 0x0008);
            this.fAutoCross = Utils.BitmaskToBool(flags, 0x0010);
            this.fLog = Utils.BitmaskToBool(flags, 0x0020);
            this.fReversed = Utils.BitmaskToBool(flags, 0x0040);
            this.fMaxCross = Utils.BitmaskToBool(flags, 0x0080);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
