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
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class ListFormatOverrideTable : List<ListFormatOverride>
    {
        private const int LFO_LENGTH = 16;
        private const int LFOLVL_LENGTH = 6;

        public ListFormatOverrideTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            if (fib.lcbPlfLfo > 0)
            {
                //read the count of LFOs
                byte[] countBytes = new byte[4];
                tableStream.Read(countBytes, 0, 4, fib.fcPlfLfo);
                Int32 count = System.BitConverter.ToInt32(countBytes, 0);

                //read the LFOs
                //int pos = fib.fcPlfLfo + 4;
                for (int i = 0; i < count; i++)
                {
                    byte[] lfoBytes = new byte[LFO_LENGTH];
                    //tableStream.Read(lfoBytes, 0, LFO_LENGTH, pos);
                    tableStream.Read(lfoBytes, LFO_LENGTH);
                    ListFormatOverride lfo = new ListFormatOverride(lfoBytes);
                    this.Add(lfo);

                    //pos += LFO_LENGTH;
                }

                //read the LFOLVLs
                for (int i = 0; i < this.Count; i++)
                {

                    for (int j = 0; j < this[i].clfolvl; j++)
                    {
                        byte[] lfolvlBytes = new byte[LFOLVL_LENGTH];
                        tableStream.Read(lfolvlBytes, LFOLVL_LENGTH);
                        this[i].rgLfoLvl[j] = new ListFormatOverrideLevel(lfolvlBytes);

                        //pos += LFOLVL_LENGTH;
                    }
                }
            }
        }
    }
}
