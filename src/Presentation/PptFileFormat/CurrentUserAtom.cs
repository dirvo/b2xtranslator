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
    [OfficeRecordAttribute(4086)]
    public class CurrentUserAtom : Record
    {
        /// <summary>
        /// Encoding to use for decoding ANSI strings.
        /// </summary>
        private static readonly Encoding ANSIEncoding = Encoding.GetEncoding("iso-8859-1");

        /// <summary>
        /// An unsigned integer that specifies the length, in bytes, of the fixed-length portion of the record. 
        /// It MUST be 0x00000014.  
        /// </summary>
        public UInt32 Size;

        /// <summary>
        /// An unsigned integer that specifies a token used to identify whether the file is encrypted.
        /// 
        /// It MUST be a value from the following table: 
        ///     0xE391C05F: The file SHOULD NOT be an encrypted document. 
        ///     0xF3D1C4DF: The file MUST be an encrypted document.
        /// </summary>
        public UInt32 HeaderToken;

        /// <summary>
        /// An unsigned integer that specifies an offset, in bytes, from the beginning of the
        /// PowerPoint DocumentRecord Stream to the UserEditAtom record for the most recent user edit. 
        /// </summary>
        public UInt32 OffsetToCurrentEdit;

        /// <summary>
        /// An unsigned integer that specifies the length, in bytes, of the  ansiUserName field.
        /// It MUST be less than or equal to 255. 
        /// </summary>
        public UInt16 LengthUserName;

        /// <summary>
        /// An unsigned integer that specifies the document file version of the file.
        /// It MUST be 0x03F4. 
        /// </summary>
        public UInt16 DocFileVersion;

        /// <summary>
        /// An unsigned integer that specifies the major version of the storage format. 
        /// It MUST be 0x03. 
        /// </summary>
        public byte MajorVersion;

        /// <summary>
        /// An unsigned integer that specifies the minor version of the storage format. 
        /// It MUST be 0x00. 
        /// </summary>
        public byte MinorVersion;

        /// <summary>
        /// A PrintableAnsiString that specifies the user name of the last user to
        /// modify the file. The length, in bytes, of the field is specified by the
        /// lenUserName field.  
        /// </summary>
        public string UserNameANSI;

        /// <summary>
        /// An unsigned integer that specifies the release version of the file format.
        /// 
        /// MUST be a value from the following table: 
        ///     0x00000008: The file contains one or more main master slide.
        ///     0x00000009: The file contains more than one main master slide. It SHOULD NOT be used. 
        /// </summary>
        public UInt32 ReleaseVersion;

        /// <summary>
        /// An optional PrintableUnicodeString that specifies the user name of the
        /// last user to modify the file.
        /// 
        /// The length, in bytes, of the field is specified by 2 * lenUserName. 
        /// 
        /// This user name supersedes that specified by the ansiUserName field.
        /// 
        /// It MAY be omitted.
        /// </summary>
        public string UserNameUnicode;

        /// <summary>
        /// UserNameUnicode if present, else UserNameANSI.
        /// </summary>
        public string UserName
        {
            get
            {
                return (this.UserNameUnicode != null) ? this.UserNameUnicode : this.UserNameANSI;
            }
        }

        public CurrentUserAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Size = this.Reader.ReadUInt32();
            this.HeaderToken = this.Reader.ReadUInt32();

            switch (this.HeaderToken)
            {
                case 0xE391C05F: /* regular PPT file */
                    break;

                case 0xF3D1C4DF: /* encrypted PPT file */
                    throw new NotSupportedException("Encryped PPT files aren't supported at this time");

                default:
                    throw new NotSupportedException(String.Format(
                        "File doesn't seem to be a PPT file. Magic Bytes = {0}", this.HeaderToken));
            }

            this.OffsetToCurrentEdit = this.Reader.ReadUInt32();
            this.LengthUserName = this.Reader.ReadUInt16();
            this.DocFileVersion = this.Reader.ReadUInt16();
            this.MajorVersion = this.Reader.ReadByte();
            this.MinorVersion = this.Reader.ReadByte();

            // Throw away reserved data
            this.Reader.ReadUInt16();

            byte[] ansiUserNameBytes = this.Reader.ReadBytes(this.LengthUserName);
            this.UserNameANSI = ANSIEncoding.GetString(ansiUserNameBytes);

            this.ReleaseVersion = this.Reader.ReadUInt32();

            if (this.Reader.BaseStream.Position != this.Reader.BaseStream.Length)
            {
                byte[] unicodeUserNameBytes = this.Reader.ReadBytes(this.LengthUserName * 2);
                this.UserNameUnicode = Encoding.Unicode.GetString(unicodeUserNameBytes);
            }
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Size = {2}, Magic = {3}, OffsetToCurrentEdit = {4}\n{1}" +
                "LengthUserName = {5}, DocFileVersion = {6}, MajorVersion = {7}, MinorVersion = {8}\n{1}" +
                "UserNameANSI = {9}, ReleaseVersion = {10}, UserNameUnicode = {11}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.Size, this.HeaderToken, this.OffsetToCurrentEdit,
                this.LengthUserName, this.DocFileVersion, this.MajorVersion, this.MinorVersion,
                this.UserNameANSI, this.ReleaseVersion, this.UserNameUnicode);
        }
    }

}
