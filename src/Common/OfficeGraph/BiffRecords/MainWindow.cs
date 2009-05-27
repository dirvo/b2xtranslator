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

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the location of the OLE server window that is contained in the parent document window when the chart data was saved.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.MainWindow)]
    public class MainWindow : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.MainWindow;

        /// <summary>
        /// A signed integer that specifies the location in twips of the left edge 
        /// of the window relative to the left edge of the primary monitor.
        /// </summary>
        public Int16 wLeft;

        /// <summary>
        /// A signed integer that specifies the location in twips of the top edge 
        /// of the window relative to the top edge of the primary monitor.
        /// </summary>
        public Int16 wTop;

        /// <summary>
        /// A signed integer that specifies the width of the window in twips.<br/>
        /// MUST be greater than or equal to 0.
        /// </summary>
        public Int16 wWidth;

        /// <summary>
        /// A signed integer that specifies the height of the window in twips.<br/>
        /// MUST be greater than or equal to 0.
        /// </summary>
        public Int16 wHeight;

        public MainWindow(IStreamReader reader, GraphRecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.wLeft = reader.ReadInt16();
            this.wTop = reader.ReadInt16();
            this.wWidth = reader.ReadInt16();
            this.wHeight = reader.ReadInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
