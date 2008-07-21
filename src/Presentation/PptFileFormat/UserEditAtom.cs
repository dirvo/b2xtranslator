/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4085)]
    public class UserEditAtom : Record
    {
        /// <summary>
        /// A SlideIdRef that specifies the last slide viewed, if this is the last 
        /// UserEditAtom record in the PowerPoint DocumentRecord Stream. In all other cases the value of this field 
        /// is undefined and MUST be ignored.
        /// </summary>
        public UInt32 LastSlideIdRef;

        /// <summary>
        /// An unsigned integer that specifies a build version of the executable that wrote the 
        /// file. It SHOULD be 0x0000 and MUST be ignored. 
        /// </summary>
        public UInt16 BuildVersion;

        /// <summary>
        /// An unsigned integer that specifies the minor version of the storage format. 
        /// It MUST be 0x00.
        /// </summary>
        public byte MinorVersion;

        /// <summary>
        /// An unsigned integer that specifies the major version of the storage format. 
        /// It MUST be 0x03. 
        /// </summary>
        public byte MajorVersion;

        /// <summary>
        /// An unsigned integer that specifies an offset, in bytes, from the beginning 
        /// of the PowerPoint DocumentRecord Stream to a UserEditAtom record for the previous user edit. It MUST 
        /// be less than the offset, in bytes, of this UserEditAtom record. The value 0x00000000 specifies that 
        /// no previous user edit exists. 
        /// </summary>
        public UInt32 OffsetLastEdit;

        /// <summary>
        /// An unsigned integer that specifies an offset, in bytes, from the 
        /// beginning of the PowerPoint DocumentRecord Stream to the PersistDirectoryAtom record for this user 
        /// edit. It MUST be greater than offsetLastEdit and less than the offset, in bytes, of this 
        /// UserEditAtom record.   
        /// </summary>
        public UInt32 OffsetPersistDirectory;

        /// <summary>
        /// A PersistIdRef that specifies the value to look up in the persist object
        /// directory to find the offset of the DocumentContainer record.
        /// It MUST be 0x00000001.
        /// </summary>
        public UInt32 DocPersistIdRef;

        /// <summary>
        /// An unsigned integer that specifies a seed for creating a new persist 
        /// object identifier. It MUST be greater than or equal to all persist object
        /// identifiers in the file as specified by the PersistDirectoryAtom records. 
        /// </summary>
        public UInt32 PersistIdSeed;

        /// <summary>
        /// A ViewTypeEnum enumeration that specifies the last view used to display the file. 
        /// </summary>
        public UInt16 LastView;

        // unused (2 bytes): Undefined and MUST be ignored. 

        /// <summary>
        /// An optional PersistIdRef that specifies the value to look up 
        /// in the persist object directory to find the offset of the CryptSession10Container record.
        /// It MAY be omitted. It MUST exist if the document is an encrypted document. 
        /// </summary>
        public UInt32? EncryptSessionPersistIdRef;

        public UserEditAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.LastSlideIdRef = this.Reader.ReadUInt32();
            this.BuildVersion = this.Reader.ReadUInt16();
            this.MinorVersion = this.Reader.ReadByte();
            this.MajorVersion = this.Reader.ReadByte();
            this.OffsetLastEdit = this.Reader.ReadUInt32();
            this.OffsetPersistDirectory = this.Reader.ReadUInt32();
            this.DocPersistIdRef = this.Reader.ReadUInt32();
            this.PersistIdSeed = this.Reader.ReadUInt32();
            this.LastView = this.Reader.ReadUInt16();

            // Throw away reserved data
            this.Reader.ReadUInt16();

            if (this.Reader.BaseStream.Position != this.Reader.BaseStream.Length)
            {
                this.EncryptSessionPersistIdRef = this.Reader.ReadUInt32();
            }
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}LastSlideIdRef = {2}, BuildVersion = {3}, MinorVersion = {4}\n{1}" +
                "MajorVersion = {5}, OffsetLastEdit = {6}, OffsetPersistDirectory = {7}, DocPersistIdRef = {8}\n{1}" +
                "PersistIdSeed = {9}, LastView = {10}, EncryptSessionPersistIdRef = {11}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.LastSlideIdRef, this.BuildVersion, this.MinorVersion,
                this.MajorVersion, this.OffsetLastEdit, this.OffsetPersistDirectory, this.DocPersistIdRef,
                this.PersistIdSeed, this.LastView, this.EncryptSessionPersistIdRef);
        }
    }

}
