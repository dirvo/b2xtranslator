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
    public class ShadingDescriptor
    {
        /// <summary>
        /// 24-bit foreground color
        /// </summary>
        public Int32 cvFore;

        /// <summary>
        /// 24-bit background color
        /// </summary>
        public Int32 cvBack;

        /// <summary>
        /// Shading pattern
        /// </summary>
        public UInt16 ipat;

        /// <summary>
        /// Creates a new ShadingDescriptor with default values
        /// </summary>
        public ShadingDescriptor()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a ShadingDescriptor.
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ShadingDescriptor(byte[] bytes)
        {
            if (bytes.Length == 10)
            {
                this.cvFore = System.BitConverter.ToInt32(bytes, 0);
                this.cvBack = System.BitConverter.ToInt32(bytes, 4);
                this.ipat = System.BitConverter.ToUInt16(bytes, 8);
            }
            else
            {
                throw new ByteParseException("SHD");
            }
        }

        private void setDefaultValues()
        {
            this.cvBack = 0;
            this.cvFore = 0;
            this.ipat = 0;
        }
    }
}
