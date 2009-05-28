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
    /// This record specifies that the plot area of the chart group is rendered in a 3-D scene 
    /// and also specifies the attributes of the 3-D plot area. 
    /// 
    /// The preceding chart group type MUST be of type bar, pie, line, area, or surface.
    /// </summary>
    [BiffRecordAttribute(RecordType.Chart3d)]
    public class Chart3d : BiffRecord
    {
        public const RecordType ID = RecordType.Chart3d;

        public enum ScalingType : ushort
        {
            /// <summary>
            /// The value of pcHeight is used to determine the height of the 3-D plot area.
            /// </summary>
            Custom = 0x0000,

            /// <summary>
            /// The height of the 3-D plot area is automatically determined
            /// </summary>
            Auto = 0x0001
        }

        /// <summary>
        /// A signed integer that specifies the clockwise rotation, in degrees, of the 3-D plot 
        /// area around a vertical line through the center of the 3-D plot area. 
        /// 
        /// MUST be greater than or equal to 0 and MUST be less than or equal to 360. 
        /// If chart group type is bar and the value of field fTranspose in the 
        /// record Bar is 1, then MUST be less than or equal to 44.
        /// </summary>
        public Int16 anRot;

        /// <summary>
        /// A signed integer that specifies the rotation, in degrees, of the 3-D plot 
        /// area around a horizontal line through the center of the 3-D plot area. 
        /// 
        /// MUST be greater than or equal to -90 and MUST be less than or equal to 90. 
        /// If the chart group type is bar and the value field fTranspose in the record Bar 
        /// is 1 or chart group type pie then MUST be greater than or equal to 0. 
        /// If the chart group type is bar and the value of field fTranspose in the 
        /// record Bar is 1, then the value MUST be less than or equal to 44.
        /// </summary>
        public Int16 anElev;

        /// <summary>
        /// A signed integer that specifies the field of view angle for the 3-D plot area. 
        /// 
        /// MUST be greater than or equal to 0 and less than or equal to 100.
        /// </summary>
        public Int16 pcDist;

        /// <summary>
        /// If fNotPieChart is 0, then this is an unsigned integer that specifies 
        /// the thickness of the pie for a pie chart group. 
        /// 
        /// If fNotPieChart is 1, then this specifies the height of the 3-D plot 
        /// area as a percentage of its width. 
        /// 
        /// MUST be greater than or equal to 5 and less than or equal to 500.
        /// </summary>
        public UInt16 pcHeight;

        /// <summary>
        /// A signed integer that specifies the depth of the 3-D plot area as a 
        /// percentage of its width. 
        /// 
        /// MUST be greater than or equal to 20 and less than or equal to 2000.
        /// </summary>
        public Int16 pcDepth;

        /// <summary>
        /// An unsigned integer that specifies the width of the gap between the series 
        /// and the front and back edges of the 3-D plot area as a percentage of the 
        /// data point depth divided by 2. If fCluster is not 1 and chart group type 
        /// is not a bar then pcGap also specifies distance between adjacent series as a 
        /// percentage of the data point depth. 
        /// 
        /// MUST be less than or equal to 500.
        /// </summary>
        public UInt16 pcGap;

        /// <summary>
        /// A bit that specifies whether the 3-D plot area is rendered with a vanishing point. 
        /// If fNotPieChart is 0 the value MUST be 0. If fNotPieChart is 1 then the 
        /// value MUST be a value from the following table: 
        /// 
        ///     Value    Meaning
        ///     0        No perspective vanishing point applied.
        ///     1        Perspective vanishing point applied based on value of pcDist.
        /// </summary>
        public bool fPerspective;

        /// <summary>
        /// A bit that specifies whether data points are clustered together in a bar chart group. 
        /// 
        /// If chart group type is not bar or pie, value MUST be ignored. 
        /// If chart group type is pie, value MUST be 0.
        /// </summary>
        public bool fCluster;

        /// <summary>
        /// A bit that specifies whether the height of the 3-D plot area is automatically determined. If fNotPieChart is 0 then this MUST be Custom.
        /// </summary>
        public ScalingType f3DScaling;

        /// <summary>
        /// A bit that specifies whether the chart group type is pie. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value     Meaning
        ///     0         Chart group type MUST be pie.
        ///     1         Chart group type MUST NOT be pie.
        /// </summary>
        public bool fNotPieChart;

        /// <summary>
        /// A bit that specifies whether the chart walls are rendered in 2-D <33>. 
        /// 
        /// If fPerspective is 1 then this MUST be ignored. 
        /// If the chart group type is not bar, area or pie this MUST be ignored. 
        /// If the chart group is of type bar and fCluster is 0, then this MUST be ignored. 
        /// If the chart group type is pie this MUST be 0 and MUST be ignored. 
        /// If the chart group type is bar or area, then the value MUST be a value from the following table:
        /// 
        ///     Value     Meaning
        ///     0         Chart walls and floor are rendered in 3D.
        ///     1         Chart walls are rendered in 2D and the chart floor is not rendered.
        /// </summary>
        public bool fWalls2D;


        public Chart3d(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.anRot = reader.ReadInt16();
            this.anElev = reader.ReadInt16();
            this.pcDist = reader.ReadInt16();
            this.pcHeight = reader.ReadUInt16();
            this.pcDepth = reader.ReadInt16();
            this.pcGap = reader.ReadUInt16();

            UInt16 flags = reader.ReadUInt16();
            this.fPerspective = Utils.BitmaskToBool(flags, 0x0001);
            this.fCluster = Utils.BitmaskToBool(flags, 0x0002);
            this.f3DScaling = (ScalingType)Utils.BitmaskToUInt16(flags, 0x0004);
            // 0x0008 is skipped
            this.fNotPieChart = Utils.BitmaskToBool(flags, 0x0010);
            this.fWalls2D = Utils.BitmaskToBool(flags, 0x0020);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
