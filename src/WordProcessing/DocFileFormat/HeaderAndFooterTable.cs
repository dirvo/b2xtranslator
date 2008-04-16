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
    public class HeaderAndFooterTable
    {
        public List<Header> FirstHeaders;
        public List<Header> EvenHeaders;
        public List<Header> OddHeaders;
        //public List<Header> FirstFooters;
        //public List<Header> EvenFooters;
        //public List<Header> OddFooters;

        public HeaderAndFooterTable(WordDocument doc)
        {
            FirstHeaders = new List<Header>();
            EvenHeaders = new List<Header>();
            OddHeaders = new List<Header>();
            //FirstFooters = new List<CharacterRange>();
            //EvenFooters = new List<CharacterRange>();
            //OddFooters = new List<CharacterRange>();
            
            //read the Table
            Int32[] table = new Int32[doc.FIB.lcbPlcfhdd / 4];
            doc.TableStream.Seek(doc.FIB.fcPlcfhdd, System.IO.SeekOrigin.Begin);
            for (int i = 0; i < table.Length; i++)
            {
                table[i] = doc.TableStream.ReadInt32();
            }

            int count = (table.Length - 8) / 6;

            //the first 6 entries are about footnote and endnote formatting
            //so skip these entries
            int pos = 6;
            for (int i = 0; i < count; i++)
            {
                Header evnHdr = new Header();
                evnHdr.CharacterPosition = doc.FIB.ccpText + table[pos];
                evnHdr.CharacterCount = table[pos + 1] - table[pos]; 
                evnHdr.Type = Header.HeaderType.EvenPage;
                this.EvenHeaders.Add(evnHdr);
                pos++;

                Header oddHdr = new Header();
                oddHdr.CharacterPosition = doc.FIB.ccpText + table[pos];
                oddHdr.CharacterCount = table[pos + 1] - table[pos];
                oddHdr.Type = Header.HeaderType.OddPage;
                this.OddHeaders.Add(oddHdr);
                pos++;

                //this.EvenFooters.Add(new CharacterRange(table[pos], table[pos + 1] - table[pos]));
                pos++;

                //this.OddFooters.Add(new CharacterRange(table[pos], table[pos + 1] - table[pos]));
                pos++;

                Header fstHdr = new Header();
                fstHdr.CharacterPosition = doc.FIB.ccpText + table[pos];
                fstHdr.CharacterCount = table[pos + 1] - table[pos];
                fstHdr.Type = Header.HeaderType.FirstPage;
                this.FirstHeaders.Add(fstHdr);
                pos++;

                //this.FirstFooters.Add(new CharacterRange(table[pos], table[pos + 1] - table[pos]));
                pos++;
            }
        }
    }
}
