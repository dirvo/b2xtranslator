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
    public class SectionTable
    {
        /// <summary>
        /// 
        /// </summary>
        public Int32[] rgfc;

        /// <summary>
        /// 
        /// </summary>
        public SectionPropertyExceptions[] grpsepx;

        private const int SEDLEN = 12;

        public SectionTable(FileInformationBlock fib, VirtualStream tableStream, VirtualStream wordStream)
        {
            byte[] bytes = new byte[fib.lcbPlcfSed];
            tableStream.Read(bytes, 0, (int)fib.lcbPlcfSed, fib.fcPlcfSed);

            //there are n SEDs and n+1 FCs
            int n = ((int)fib.lcbPlcfSed - 4) / (4 + SEDLEN);
            this.grpsepx = new SectionPropertyExceptions[n];
            this.rgfc = new Int32[n + 1];

            int pos = 0;

            //read the FCs
            for (int i = 0; i < n+1; i++)
            {
                this.rgfc[i] = System.BitConverter.ToInt32(bytes, pos);
                pos += 4;
            }

            //read the SEDS
            for (int i = 0; i < n ; i++)
            {
                //skip the first 2 bytes of the SED
                pos += 2;

                //read the FC of the SEPX
                int fcSepx = System.BitConverter.ToInt32(bytes, pos);

                //read the cb of the SEPX
                byte[] cbBytes = new byte[2];
                wordStream.Read(cbBytes, 0, 2, fcSepx);
                Int16 cbSepx = System.BitConverter.ToInt16(cbBytes, 0);
                pos += 4;

                if (cbSepx > 0)
                {
                    //parse the SEPX and add it to the list
                    byte[] sepxBytes = new byte[cbSepx - 2];
                    wordStream.Read(sepxBytes, 0, sepxBytes.Length, fcSepx + 2);
                    this.grpsepx[i] = new SectionPropertyExceptions(sepxBytes);
                }

                //skip the last 6 bytes of the SED
                pos += 6;
            }
        }
    }
}
