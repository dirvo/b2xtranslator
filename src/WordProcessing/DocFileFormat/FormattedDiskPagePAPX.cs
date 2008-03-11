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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FormattedDiskPagePAPX : FormattedDiskPage
    {
        /// <summary>
        /// An array of the BX data structure.<br/>
        /// BX is a 13 byte data structure. The first byte of each is the word offset of the PAPX.
        /// </summary>
        public BX[] rgbx;

        /// <summary>
        /// grppapx consists of all of the PAPXs stored in FKP concatenated end to end. 
        /// Each PAPX begins with a count of words which records its length padded to a word boundary.
        /// </summary>
        public ParagraphPropertyExceptions[] grppapx;

        public struct BX
        {
            public byte wordOffset;
            public ParagraphHeight phe;
        }

        public FormattedDiskPagePAPX(VirtualStream wordStream, Int32 offset)
        {
            this.Type = FKPType.Paragraph;
            this.WordStream = wordStream;
            this.Offset = offset;

            //read the 512 bytes (FKP)
            byte[] bytes = new byte[512];
            wordStream.Read(bytes, 512, offset);

            //get the count first
            this.crun = bytes[511];

            //create and fill the array with the adresses
            this.rgfc = new Int32[this.crun + 1];
            int j = 0;
            for (int i = 0; i < rgfc.Length; i++)
            {
                rgfc[i] = System.BitConverter.ToInt32(bytes, j);
                j += 4;
            }

            //create arrays
            this.rgbx = new BX[this.crun];
            this.grppapx = new ParagraphPropertyExceptions[this.crun];

            j = 4*(this.crun+1);
            for (int i = 0; i < rgbx.Length; i++)
            {
                //read the 12 for PHE
                byte[] phe = new byte[12];
                Array.Copy(bytes, j+1, phe, 0, phe.Length);

                //fill the rgbx array
                BX bx = new BX();
                bx.wordOffset = bytes[j];
                bx.phe = new ParagraphHeight(phe, false);
                rgbx[i] = bx;
                j += 13;

                if (bx.wordOffset != 0)
                {
                    //read first byte of PAPX
                    //PAPX is stored in a FKP; so the first byte is a count of words
                    byte padbyte = 0;
                    byte cw = bytes[bx.wordOffset * 2];

                    //if that byte is zero, it's a pad byte, and the word count is the following byte
                    if (cw == 0)
                    {
                        padbyte = 1;
                        cw = bytes[bx.wordOffset * 2 + 1];
                    }

                    if (cw != 0)
                    {
                        //read the bytes for papx
                        byte[] papx = new byte[cw * 2];
                        Array.Copy(bytes, (bx.wordOffset * 2) + padbyte + 1, papx, 0, papx.Length);

                        //parse PAPX and fill grppapx
                        this.grppapx[i] = new ParagraphPropertyExceptions(papx);
                    }
                }
                else
                {
                    //create a PAPX which doesn't modify anything
                    this.grppapx[i] = new ParagraphPropertyExceptions();
                }
            }
        }

        /// <summary>
        /// Parses the 0Table (or 1Table) for FKP entries containing PAPX
        /// </summary>
        /// <param name="fib">The FileInformationBlock</param>
        /// <param name="wordStream">The WordDocument stream</param>
        /// <param name="tableStream">The 0Table stream</param>
        /// <returns></returns>
        public static List<FormattedDiskPagePAPX> GetAllPAPXFKPs(FileInformationBlock fib, VirtualStream wordStream, VirtualStream tableStream)
        {
            List<FormattedDiskPagePAPX> list = new List<FormattedDiskPagePAPX>();

            //get bintable for PAPX
            byte[] binTablePapx = new byte[fib.lcbPlcfbtePapx];
            tableStream.Read(binTablePapx, binTablePapx.Length, (int)fib.fcPlcfbtePapx);

            //there are n offsets and n-1 fkp's in the bin table
            int n = (((int)fib.lcbPlcfbtePapx - 4) / 8) + 1;

            //Get the indexed PAPX FKPs
            for (int i = (n * 4); i < binTablePapx.Length; i += 4)
            {
                //indexed FKP is the xth 512byte page
                int fkpnr = System.BitConverter.ToInt32(binTablePapx, i);

                //so starts at:
                int offset = fkpnr * 512;

                //parse the FKP and add it to the list
                list.Add(new FormattedDiskPagePAPX(wordStream, offset));
            }

            return list;
        }

        public static List<Int32> GetAllFCs(FileInformationBlock fib, VirtualStream wordStream, VirtualStream tableStream)
        {
            List<Int32> list = new List<Int32>();

            //get bintable for PAPX
            byte[] binTablePapx = new byte[fib.lcbPlcfbtePapx];
            tableStream.Read(binTablePapx, binTablePapx.Length, (int)fib.fcPlcfbtePapx);

            //there are n offsets and n-1 fkp's in the bin table
            int n = (((int)fib.lcbPlcfbtePapx - 4) / 8) + 1;

            //Get the indexed PAPX FKPs
            for (int i = (n * 4); i < binTablePapx.Length; i += 4)
            {
                //indexed FKP is the xth 512byte page
                int fkpnr = System.BitConverter.ToInt32(binTablePapx, i);

                //so starts at:
                int offset = fkpnr * 512;

                //parse the FKP and add offset to the list
                FormattedDiskPagePAPX fkp = new FormattedDiskPagePAPX(wordStream, offset);
                foreach (int fc in fkp.rgfc)
                {
                    //don't add the duplicated values of the FKP boundaries
                    if(!list.Contains(fc))
                        list.Add(fc);
                }
            }

            return list;
        }

        /// <summary>
        /// Returnes a list of all ParagraphPropertyExceptions which correspond to text 
        /// between the given offsets.
        /// </summary>
        /// <param name="fcStart"></param>
        /// <param name="fcEnd"></param>
        /// <param name="fib"></param>
        /// <param name="wordStream"></param>
        /// <param name="tableStream"></param>
        /// <returns></returns>
        public static List<ParagraphPropertyExceptions> GetAllPAPX(
            Int32 fcStart,
            Int32 fcEnd,
            FileInformationBlock fib, 
            VirtualStream wordStream, 
            VirtualStream tableStream)
        {
            List<ParagraphPropertyExceptions> list = new List<ParagraphPropertyExceptions>();

            //get bintable for PAPX
            byte[] binTablePapx = new byte[fib.lcbPlcfbtePapx];
            tableStream.Read(binTablePapx, binTablePapx.Length, (int)fib.fcPlcfbtePapx);

            //there are n offsets and n-1 fkp's in the bin table
            int n = (((int)fib.lcbPlcfbtePapx - 4) / 8) + 1;

            //Get the indexed PAPX FKPs
            for (int i = (n * 4); i < binTablePapx.Length; i += 4)
            {
                //indexed FKP is the xth 512byte page
                int fkpnr = System.BitConverter.ToInt32(binTablePapx, i);

                //so starts at:
                int offset = fkpnr * 512;

                //parse the FKP and add PAPX to the list
                FormattedDiskPagePAPX fkp = new FormattedDiskPagePAPX(wordStream, offset);
                for (int j = 0; j < fkp.grppapx.Length; j++)
                {
                    if (fkp.rgfc[j] >= fcStart && fkp.rgfc[j] < fcEnd)
                    {
                        list.Add(fkp.grppapx[j]);
                    }
                }
            }

            return list;
        }

    }
}
