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
using System.Collections;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FontFamilyName
    {
        public struct FontSignature
        {
            public UInt32 UnicodeSubsetBitfield0;
            public UInt32 UnicodeSubsetBitfield1;
            public UInt32 UnicodeSubsetBitfield2;
            public UInt32 UnicodeSubsetBitfield3;
            public UInt32 CodePageBitfield0;
            public UInt32 CodePageBitfield1;
        }

        /// <summary>
        /// When true, font is a TrueType font
        /// </summary>
        public bool fTrueType;

        /// <summary>
        /// Font family id
        /// </summary>
        public byte ff;

        /// <summary>
        /// Base weight of font
        /// </summary>
        public Int16 wWeight;

        /// <summary>
        /// Character set identifier
        /// </summary>
        public byte chs;

        /// <summary>
        /// Pitch request
        /// </summary>
        public byte prq;

        /// <summary>
        /// Name of font
        /// </summary>
        public String xszFtn;

        /// <summary>
        /// Panose
        /// </summary>
        public byte[] panose;

        /// <summary>
        /// Font sinature
        /// </summary>
        public FontSignature fs;

        /// <summary>
        /// Parses the byte to retrieve a FFN structure
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public FontFamilyName(byte[] bytes)
        {
            if (bytes.Length > 40)
            {
                int cbFfnM1 = bytes[0];

                //unmask byte 1
                int req = (int)bytes[1];
                req = req << 6;
                req = req >> 6;
                this.prq = (byte)req;

                this.fTrueType = Utils.BitmaskToBool((int)bytes[1], 0x04);

                int family = (int)bytes[1];
                family = family << 1;
                family = family >> 4;
                this.ff = (byte)family;

                //byte 2 and 3
                this.wWeight = System.BitConverter.ToInt16(bytes, 2);

                //byte 4
                this.chs = bytes[4];

                //byte 5

                //byte 6-15
                this.panose = new byte[10];
                Array.Copy(bytes, 6, panose, 0, 10);

                //byte 16 - 39
                this.fs = new FontSignature();
                this.fs.UnicodeSubsetBitfield0 = System.BitConverter.ToUInt32(bytes, 16);
                this.fs.UnicodeSubsetBitfield1 = System.BitConverter.ToUInt32(bytes, 20);
                this.fs.UnicodeSubsetBitfield2 = System.BitConverter.ToUInt32(bytes, 24);
                this.fs.UnicodeSubsetBitfield3 = System.BitConverter.ToUInt32(bytes, 28);
                this.fs.CodePageBitfield0 = System.BitConverter.ToUInt32(bytes, 32);
                this.fs.CodePageBitfield1 = System.BitConverter.ToUInt32(bytes, 36);

                //byte 40 - x (name)
                byte[] name = new byte[bytes.Length - 40];
                Array.Copy(bytes, 40, name, 0, name.Length);
                this.xszFtn = Encoding.Unicode.GetString(name);
                this.xszFtn = this.xszFtn.Replace("\0", "");
            }
        }
    }
}
