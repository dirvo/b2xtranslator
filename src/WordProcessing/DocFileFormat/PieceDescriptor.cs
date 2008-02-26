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

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class PieceDescriptor
    {
        /// <summary>
        /// File offset of beginning of piece. <br/>
        /// This is relative to the beginning of the WordDocument stream.
        /// </summary>
        public UInt16 fc;

        /// <summary>
        /// Character start position of this piece.<br/>
        /// </summary>
        public Int32 cpStart;

        /// <summary>
        /// Character end position of this piece<br/>
        /// </summary>
        public Int32 cpEnd;

        /// <summary>
        /// The encoding of the piece
        /// </summary>
        public Encoding encoding;

        /// <summary>
        /// Parses the bytes to retrieve a PieceDescriptor
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public PieceDescriptor(byte[] bytes)
        {
            int encodingMask = (int)System.BitConverter.ToUInt16(bytes, 4);

            //find encoding and offset
            encodingMask = encodingMask >> 14;
            if (encodingMask == 1)
            {
                this.encoding = Encoding.GetEncoding(1252);
                this.fc = (UInt16)(System.BitConverter.ToUInt16(bytes, 2) / 2);
            }
            else
            {
                this.encoding = Encoding.Unicode;
                this.fc = System.BitConverter.ToUInt16(bytes, 2);
            }
        }
    }
}
