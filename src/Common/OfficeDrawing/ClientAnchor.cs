
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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(0xF010)]
    public class ClientAnchor : Record
    {
        public byte[] Bytes;
        public int Top;
        public int Left;
        public int Right;
        public int Bottom;

        public ClientAnchor(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        {
            this.Bytes = this.Reader.ReadBytes((int)this.BodySize);

            this.Top = System.BitConverter.ToInt16(this.Bytes, 0); // this.Reader.ReadUInt16();
            this.Left = System.BitConverter.ToInt16(this.Bytes, 2); //this.Reader.ReadUInt16();
            this.Right = System.BitConverter.ToInt16(this.Bytes, 4); //this.Reader.ReadUInt16();
            this.Bottom = System.BitConverter.ToInt16(this.Bytes, 6); //this.Reader.ReadUInt16();
        }        

        public override string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Top = {2}, Left = {3}, Bottom = {4}, Right = {5})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.Top, this.Left, this.Bottom, this.Right);
        }

    }

}
