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

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FontFamilyName
    {
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
        /// Name of font
        /// </summary>
        public String xszFtn;

        /// <summary>
        /// Parses the byte to retrieve a FFN structure
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public FontFamilyName(byte[] bytes)
        {
            if (bytes.Length > 40)
            {
                this.wWeight = System.BitConverter.ToInt16(bytes, 2);
                this.chs = bytes[4];

                //copy name to array
                byte[] name = new byte[bytes.Length - 40];
                Array.Copy(bytes, 40, name, 0, name.Length);
                this.xszFtn = Encoding.Unicode.GetString(name);

                //replace zero termination
                this.xszFtn = this.xszFtn.Replace("\0", "");
            }
        }
    }
}
