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
    public class PieceTable
    {
        /// <summary>
        /// A list of PieceDescriptor standing for each piece of text.
        /// </summary>
        public List<PieceDescriptor> Pieces;

        /// <summary>
        /// Parses the pice table and creates a list of PieceDescriptors.
        /// </summary>
        /// <param name="fib">The FIB</param>
        /// <param name="tableStream">The 0Table or 1Table stream</param>
        public PieceTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            //Read the bytes of complex file information
            byte[] bytes = new byte[fib.lcbClx];
            tableStream.Read(bytes, (int)fib.lcbClx, (int)fib.fcClx);

            this.Pieces = new List<PieceDescriptor>();

            int pos = 0;
            bool goon = true;
            while (goon)
            {
                try
                {
                    byte type = bytes[pos];

                    //check if the type of the entry is a piece table
                    if (type == 2)
                    {
                        Int32 lcb = System.BitConverter.ToInt32(bytes, pos + 1);

                        //read the piece table
                        byte[] piecetable = new byte[lcb];
                        Array.Copy(bytes, pos + 5, piecetable, 0, piecetable.Length);

                        //count of entries
                        int n = (lcb - 4) / 12;

                        //and n piece descriptors
                        for (int i = 0; i < n; i++)
                        {
                            byte[] pcdBytes = new byte[8];
                            int index = ((n+1)*4) + (i*8);
                            Array.Copy(piecetable, index, pcdBytes, 0, 8);

                            //build pcd
                            PieceDescriptor pcd = new PieceDescriptor(pcdBytes);
                            pcd.cpStart = System.BitConverter.ToInt32(piecetable, i * 4);
                            pcd.cpEnd = System.BitConverter.ToInt32(piecetable, (i+1) * 4);

                            this.Pieces.Add(pcd);
                        }

                        //piecetable was found
                        goon = false;
                    }
                    //entry is no piecetable so goon
                    else if (type == 1)
                    {
                        byte cb = bytes[pos + 1];
                        pos = pos + 1 + 1 + cb;
                    }
                    else
                    {
                        goon = false;
                    }
                }
                catch(Exception)
                {
                    goon = false;
                }
            }
        }
    }
}
