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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class SinglePropertyModifier 
    {
        /// <summary>
        /// Identifies the type of a SPRM
        /// </summary>
        public enum SprmType
        {
            PAP = 1,
            CHP = 2,
            PIC = 3,
            SEP = 4,
            TAP = 5
        }

        /// <summary>
        /// The operation code identifies the property of the 
        /// PAP/CHP/PIC/SEP/TAP which sould be modified
        /// </summary>
        public UInt16 OpCode;

        /// <summary>
        /// This SPRM requires special handling
        /// </summary>
        public bool fSpec;

        /// <summary>
        /// The type of the SPRM
        /// </summary>
        public SprmType Type;

        /// <summary>
        /// The arguments which is applied to the property
        /// </summary>
        public byte[] Arguments;

        /// <summary>
        /// parses the byte to retrieve a SPRM
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public SinglePropertyModifier(byte[] bytes)
        {
            //first 2 bytes are the operation code ...
            this.OpCode = System.BitConverter.ToUInt16(bytes, 0);

            //... whereof bit 9 is fSpec ...
            UInt32 j = (UInt32)this.OpCode << 22;
            j = j >> 31;
            if (j == 1)
                this.fSpec = true;
            else
                this.fSpec = false;

            //... and bits 10,11,12 are the type ...
            UInt32 i = (UInt32)this.OpCode << 19;
            i = i >> 29;
            this.Type = (SprmType)i;

            //... and last 3 bits are the spra
            byte spra = (byte)((Int32)this.OpCode >> 13);
            byte opSize = GetOperandSize(spra);
            if (opSize == 255)
            {
                //the variable length stand in the byte after the opcode (byte2)
                opSize = bytes[2];
                //and the arguments start at the byte after that (byte3)
                this.Arguments = new byte[opSize];
                Array.Copy(bytes, 3, this.Arguments, 0, Arguments.Length);
            }
            else
            {
                this.Arguments = new byte[opSize];
                Array.Copy(bytes, 2, this.Arguments, 0, Arguments.Length);
            }
        }

        /// <summary>
        /// Get be used to get the size of the sprm's operand.
        /// Returns 0 if the Operation failed and 255 if the size is variable
        /// </summary>
        /// <param name="spra">the 3 bits for spra (as byte)</param>
        /// <returns>the size (as byte)</returns>
        public static byte GetOperandSize(byte spra)
        {
            switch (spra)
            {
                case 0: return 1;
                case 1: return 1;
                case 2: return 2;
                case 3: return 4;
                case 4: return 2;
                case 5: return 2;
                case 6: return 255;
                case 7: return 3;
                default: return 0;
            }
        }
    }
}
