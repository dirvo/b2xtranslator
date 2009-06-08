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
    /// This structure specifies the common properties of the Obj record that contains this FtCmo.
    /// </summary>
    public class FtCmo
    {
        public enum ObjectType : ushort
        {
            Group = 0x0000,
            Line = 0x0001,
            Rectangle = 0x0002,
            Oval = 0x0003,
            Arc = 0x0004,
            Chart = 0x0005,
            Text = 0x0006,
            Button = 0x0007,
            Picture = 0x0008,
            Polygon = 0x0009,
            Checkbox = 0x000B,
            RadioButton = 0x000C,
            EditBox = 0x000D,
            Label = 0x000E,
            DialogBox = 0x000F,
            SpinControl = 0x0010,
            Scrollbar = 0x0011,
            List = 0x0012,
            GroupBox = 0x0013,
            DropdownList = 0x0014,
            Note = 0x0019,
            OfficeArtObject = 0x001E
        }

        /// <summary>
        /// Reserved. MUST be 0x15.
        /// </summary>
        public UInt16 ft;

        /// <summary>
        /// Reserved. MUST be 0x12.
        /// </summary>
        public UInt16 cb;

        /// <summary>
        /// An unsigned integer that specifies the type of object represented by the Obj record that contains this FtCmo
        /// </summary>
        public ObjectType ot;

        /// <summary>
        /// An unsigned integer that specifies the identifier of this object. This object identifier 
        /// is used by other types to refer to this object. 
        /// 
        /// The value of id MUST be unique among all Obj records within the Chart Sheet 
        /// Substream ABNF, Macro Sheet Substream ABNF and Worksheet Substream ABNF.
        /// </summary>
        public UInt16 id;

        /// <summary>
        /// A bit that specifies whether this object is locked.
        /// </summary>
        public bool fLocked;

        /// <summary>
        /// A bit that specifies whether the application is expected to choose the object‘s size.
        /// </summary>
        public bool fDefaultSize;

        /// <summary>
        /// A bit that specifies whether this is a chart object that is expected to be published 
        /// the next time the sheet containing it is published <158>. This bit is ignored if the 
        /// fPublishedBookItems field of the BookExt_Conditional12 structure is zero.
        /// </summary>
        public bool fPublished;

        /// <summary>
        /// A bit that specifies whether the image of this object is intended to be included when printed.
        /// </summary>
        public bool fPrint;

        /// <summary>
        /// A bit that specifies whether this object has been disabled.
        /// </summary>
        public bool fDisabled;

        /// <summary>
        /// A bit that specifies whether this is an auxiliary object that can only be automatically 
        /// inserted by the application (as opposed to an object that can be inserted by a user).
        /// </summary>
        public bool fUIObj;

        /// <summary>
        /// A bit that specifies whether this object is expected to be updated on load to 
        /// reflect the values in the range associated with the object. 
        /// 
        /// This field MUST be ignored unless the pictfmla.key field of the containing Obj 
        /// exists and pictfmla.key.fmlaListFillRange.cbFmla of the containing Obj is not equal to 0.
        /// </summary>
        public bool fRecalcObj;

        /// <summary>
        /// A bit that specifies whether this object is expected to be updated whenever 
        /// the value of a cell in the range associated with the object changes. 
        /// 
        /// This field MUST be ignored unless the pictfmla.key field of the containing Obj 
        /// exists and pictfmla.key.fmlaListFillRange.cbFmla of the containing Obj is not equal to 0.
        /// </summary>
        public bool fRecalcObjAlways;

        public FtCmo(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();
            this.ot = (ObjectType)reader.ReadUInt16();
            this.id = reader.ReadUInt16();

            UInt16 flags = reader.ReadUInt16();
            this.fLocked = Utils.BitmaskToBool(flags, 0x0001);

            this.fDefaultSize = Utils.BitmaskToBool(flags, 0x0004);
            this.fPublished = Utils.BitmaskToBool(flags, 0x0008);
            this.fPrint = Utils.BitmaskToBool(flags, 0x0010);

            this.fDisabled = Utils.BitmaskToBool(flags, 0x0080);
            this.fUIObj = Utils.BitmaskToBool(flags, 0x0100);
            this.fRecalcObj = Utils.BitmaskToBool(flags, 0x0200);

            this.fRecalcObjAlways = Utils.BitmaskToBool(flags, 0x1000);

            reader.ReadBytes(12);
        }
    }
}

