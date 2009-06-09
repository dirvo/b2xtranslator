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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies information about the embedded control associated with the 
    /// Obj record that contains the ObjFmla that contains this PictFmlaEmbedInfo. 
    /// The embedded control can be an ActiveX control, an OLE object or a camera picture control.
    /// The pictFlags field of this Obj record specifies the type of embedded control.
    /// </summary>
    public class PictFmlaEmbedInfo
    {
        /// <summary>
        /// Reserved. MUST be 0x03.
        /// </summary>
        public byte ttb;

        /// <summary>
        /// An unsigned integer that specifies the length in bytes of the strClass field.
        /// </summary>
        public byte cbClass;

        /// <summary>
        /// An optional XLUnicodeStringNoCch that specifies the class name of the embedded 
        /// control associated with this Obj. This field MUST exist if and only if cbClass is nonzero.
        /// </summary>
        public XLUnicodeStringNoCch strClass;


        public PictFmlaEmbedInfo(IStreamReader reader)
        {
            this.ttb = reader.ReadByte();
            this.cbClass = reader.ReadByte();
            reader.ReadByte();

            if (this.cbClass > 0)
            {
                this.strClass = new XLUnicodeStringNoCch(reader, this.cbClass);
            }
        }
    }
}