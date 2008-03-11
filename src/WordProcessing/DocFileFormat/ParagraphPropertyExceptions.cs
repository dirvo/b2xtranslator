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
    public class ParagraphPropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Index to style descriptor of the style from which the 
        /// paragraph inherits its paragraph and character properties
        /// </summary>
        public UInt16 istd;

        /// <summary>
        /// Creates a PAPX wich doesn't modify anything.<br/>
        /// The grpprl list is null
        /// </summary>
        public ParagraphPropertyExceptions()
        {
            grpprl = new List<SinglePropertyModifier>();
        }

        /// <summary>
        /// Parses the bytes to retrieve a PAPX
        /// </summary>
        /// <param name="bytes">The bytes starting with the istd</param>
        public ParagraphPropertyExceptions(byte[] bytes)
        {
            if (bytes.Length != 0)
            {
                this.istd = System.BitConverter.ToUInt16(bytes, 0);

                //read the sprms
                grpprl = new List<SinglePropertyModifier>();
                int sprmStart = 2;
                bool goOn = true;
                while (goOn)
                {
                    //enough bytes to read?
                    if(sprmStart+2 < bytes.Length)
                    {
                        //make spra
                        UInt16 opCode = System.BitConverter.ToUInt16(bytes, sprmStart);
                        byte spra = (byte)((Int32)opCode >> 13);

                        // get size of operand
                        byte opSize = SinglePropertyModifier.GetOperandSize(spra);
                        byte lenByte = 0;
                        if (opSize == 255)
                        {
                            //the variable length stand in the byte after the opcode
                            lenByte = 1;
                            opSize = bytes[sprmStart + 2];
                        }

                        //copy sprm to array
                        byte[] sprm = new byte[2 + lenByte + opSize];

                        if (bytes.Length >= sprmStart + sprm.Length)
                        {
                            Array.Copy(bytes, sprmStart, sprm, 0, sprm.Length);
                            //parse and save
                            grpprl.Add(new SinglePropertyModifier(sprm));

                            sprmStart += sprm.Length;
                        }
                        else
                        {
                            goOn = false;
                        }
                    }
                    else
                    {
                        goOn = false;
                    }
                }
                
            }
            else
            {
                throw new ByteParseException("PAPX");
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ParagraphPropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }
}
