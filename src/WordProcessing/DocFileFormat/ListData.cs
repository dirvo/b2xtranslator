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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class ListData
    {
        /// <summary>
        /// Unique List ID
        /// </summary>
        public Int32 lsid;

        /// <summary>
        /// Unique template code
        /// </summary>
        public Int32 tplc;

        /// <summary>
        /// Array of shorts containing the istd�s linked to each level of the list, 
        /// or ISTD_NIL (4095) if no style is linked.
        /// </summary>
        public Int16[] rgistd;

        /// <summary>
        /// True if this is a simple (one-level) list.<br/>
        /// False if this is a multilevel (nine-level) list.
        /// </summary>
        public bool fSimpleList;

        /// <summary>
        /// Word 6.0 compatibility option:<br/>
        /// True if the list should start numbering over at the beginning of each section.
        /// </summary>
        public bool fRestartHdn;

        /// <summary>
        /// To emulate Word 6.0 numbering: <br/>
        /// True if Auto numbering
        /// </summary>
        public bool fAutoNum;

        /// <summary>
        /// When true, this list was there before we started reading RTF.
        /// </summary>
        public bool fPreRTF;

        /// <summary>
        /// When true, list is a hybrid multilevel/simple (UI=simple, internal=multilevel)
        /// </summary>
        public bool fHybrid;

        /// <summary>
        /// Array of ListLevel describing the several levels of the list.
        /// </summary>
        public ListLevel[] rglvl;

        public const Int16 ISTD_NIL = 4095;
        private const int RGISTD_LENGTH = 9;

        /// <summary>
        /// Parses the bytes to retrieve a ListData
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ListData(byte[] bytes)
        {
            if (bytes.Length == 28)
            {
                this.lsid = System.BitConverter.ToInt32(bytes, 0);
                this.tplc = System.BitConverter.ToInt32(bytes, 4);

                this.rgistd = new Int16[RGISTD_LENGTH];
                for (int i = 0; i < RGISTD_LENGTH; i++)
                {
                    this.rgistd[i] = System.BitConverter.ToInt16(bytes, 8 + (i*2));
                }

                //parse flagbyte
                int flag = (int)bytes[26];
                this.fSimpleList = Utils.BitmaskToBool(flag, 0x01);

                if(this.fSimpleList)
                    this.rglvl = new ListLevel[1];
                else
                    this.rglvl = new ListLevel[9];

                this.fRestartHdn = Utils.BitmaskToBool(flag, 0x02);
                this.fAutoNum = Utils.BitmaskToBool(flag, 0x04);
                this.fPreRTF = Utils.BitmaskToBool(flag, 0x08);
                this.fHybrid = Utils.BitmaskToBool(flag, 0x10);
            }
            else
            {
                throw new ByteParseException("LSTF");
            }
        }
    }
}
