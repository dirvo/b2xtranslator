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
    public class BorderCode
    {
        /// <summary>
        /// 24-bit border color
        /// </summary>
        public Int32 cv;

        /// <summary>
        /// Width of a single line in 1/8pt, max of 32pt
        /// </summary>
        public Int32 dptLineWidth;

        /// <summary>
        /// Border type code:
        /// 0 none
        /// 1 single
        /// 2 thick
        /// 3 double
        /// 5 hairline
        /// 6 dot
        /// 7 dash large gap
        /// 8 dot dash
        /// 9 dot dot dash
        /// 10 triple
        /// 11 thin-thick small gap
        /// 12 tick-thin small gap
        /// 13 thin-thick-thin small gap
        /// 14 thin-thick medium gap
        /// 15 thick-thin medium gap
        /// 16 thin-thick-thin medium gap
        /// 17 thin-thick large gap
        /// 18 thick-thin large gap
        /// 19 thin-thick-thin large gap
        /// 20 wave
        /// 21 double wave
        /// 22 dash small gap
        /// 23 dash dot stroked
        /// 24 emboss 3D
        /// 25 engrave 3D
        /// </summary>
        public Int32 brcType;

        /// <summary>
        /// The color of the Border.
        /// Unused if cv is set.
        /// </summary>
        public Color.ColorIdentifier ico;

        /// <summary>
        /// Width of space to maintain between border and text within border
        /// </summary>
        public Int32 dptSpace;

        /// <summary>
        /// When true, border is drawn with shadow. Must be false when BRC is substructure of the TC
        /// </summary>
        public bool fShadow;

        /// <summary>
        /// When true, don't reverse the border
        /// </summary>
        public bool fFrame;

        /// <summary>
        /// Creates a new BorderCode with default values
        /// </summary>
        public BorderCode()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the byte for a BRC
        /// </summary>
        /// <param name="bytes"></param>
        public BorderCode(byte[] bytes)
        {
            if (bytes.Length == 8)
            {
                //it's a border code of Word 2000/2003
                this.cv = System.BitConverter.ToInt32(bytes, 0);
                this.ico = Color.ColorIdentifier.Auto;
                Int32 val = System.BitConverter.ToInt32(bytes, 4);
                this.dptLineWidth = val & 0x000000FF;
                this.brcType = val & 0x0000FF00;
                this.dptSpace = val & 0x001F0000;
                this.fShadow = Utils.BitmaskToBool(val, 0x00200000);
                this.fFrame = Utils.BitmaskToBool(val, 0x00400000);
            }
            else if (bytes.Length == 4)
            {
                //it's a border code of Word 97
                UInt16 val = System.BitConverter.ToUInt16(bytes, 0);
                this.dptLineWidth = val & 0x00FF;
                this.brcType = (val & 0xFF00) >> 8;
                val = System.BitConverter.ToUInt16(bytes, 2);
                this.ico = (Color.ColorIdentifier)(val & 0x00FF);
                this.dptSpace = (val & 0x1F00) >> 8;
            }
            else
            {
                throw new ByteParseException("BRC");
            }
        }

        private void setDefaultValues()
        {
            this.brcType = 0;
            this.cv = 0;
            this.ico = Color.ColorIdentifier.Auto;
            this.dptLineWidth = 0;
            this.dptSpace = 0;
            this.fFrame = false;
            this.fShadow = false;
        }
    }
}
