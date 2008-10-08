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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class PieceTable
    {
        /// <summary>
        /// A list of PieceDescriptor standing for each piece of text.
        /// </summary>
        public List<PieceDescriptor> Pieces;

        /// <summary>
        /// A dictionary with character positions as keys and the matching FCs as values
        /// </summary>
        public Dictionary<Int32, Int32> FileCharacterPositions;

        /// <summary>
        /// A dictionary with file character positions as keys and the matching CPs as values
        /// </summary>
        public Dictionary<Int32, Int32> CharacterPositions;

        /// <summary>
        /// Parses the pice table and creates a list of PieceDescriptors.
        /// </summary>
        /// <param name="fib">The FIB</param>
        /// <param name="tableStream">The 0Table or 1Table stream</param>
        public PieceTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            //Read the bytes of complex file information
            byte[] bytes = new byte[fib.lcbClx];
            tableStream.Read(bytes, 0, (int)fib.lcbClx, (int)fib.fcClx);

            this.Pieces = new List<PieceDescriptor>();
            this.FileCharacterPositions = new Dictionary<int, int>();
            this.CharacterPositions = new Dictionary<int, int>();

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

                        //count of PCD _entries
                        int n = (lcb - 4) / 12;

                        //and n piece descriptors
                        for (int i = 0; i < n; i++)
                        {
                            //read the CP 
                            int indexCp = i * 4;
                            Int32 cp = System.BitConverter.ToInt32(piecetable, indexCp);

                            //read the next CP
                            int indexCpNext = (i+1) * 4;
                            Int32 cpNext = System.BitConverter.ToInt32(piecetable, indexCpNext);

                            //read the PCD
                            int indexPcd = ((n + 1) * 4) + (i * 8);
                            byte[] pcdBytes = new byte[8];
                            Array.Copy(piecetable, indexPcd, pcdBytes, 0, 8);
                            PieceDescriptor pcd = new PieceDescriptor(pcdBytes);
                            pcd.cpStart = cp;
                            pcd.cpEnd = cpNext;
                            
                            //add pcd
                            this.Pieces.Add(pcd);

                            //add positions
                            Int32 f = (Int32)pcd.fc;
                            Int32 multi = 1;
                            if (pcd.encoding == Encoding.Unicode)
                            {
                                multi = 2;
                            }
                            for (int c = pcd.cpStart; c < pcd.cpEnd; c++)
                            {
                                if (!this.FileCharacterPositions.ContainsKey(c))
                                    this.FileCharacterPositions.Add(c, f);
                                if (!this.CharacterPositions.ContainsKey(f))
                                    this.CharacterPositions.Add(f, c);

                                f += multi;
                            }
                        }
                        this.FileCharacterPositions.Add(this.FileCharacterPositions.Count, fib.fcMac);
                        this.CharacterPositions.Add(fib.fcMac, this.FileCharacterPositions.Count);

                        //piecetable was found
                        goon = false;
                    }
                    //entry is no piecetable so goon
                    else if (type == 1)
                    {
                        Int16 cb = System.BitConverter.ToInt16(bytes, pos + 1);
                        pos = pos + 1 + 2 + cb;
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

        public List<char> GetAllChars(VirtualStream wordStream)
        {
            List<char> chars = new List<char>();
            foreach (PieceDescriptor pcd in this.Pieces)
            {
                //get the FC end of this piece
                Int32 pcdFcEnd = pcd.cpEnd - pcd.cpStart;
                if (pcd.encoding == Encoding.Unicode)
                    pcdFcEnd *= 2;
                pcdFcEnd += (Int32)pcd.fc;

                int cb = pcdFcEnd - (Int32)pcd.fc;
                byte[] bytes = new byte[cb];

                //read all bytes 
                wordStream.Read(bytes, 0, cb, (Int32)pcd.fc);

                //get the chars
                char[] plainChars = pcd.encoding.GetString(bytes).ToCharArray();

                //add to list
                foreach (char c in plainChars)
                {
                    chars.Add(c);
                }
            }
            return chars;
        }

        public List<char> GetChars(Int32 fcStart, Int32 fcEnd, VirtualStream wordStream)
        {
            List<char> chars = new List<char>();
            for (int i = 0; i < this.Pieces.Count; i++)
            {
                PieceDescriptor pcd = this.Pieces[i];

                //get the FC end of this piece
                Int32 pcdFcEnd = pcd.cpEnd - pcd.cpStart;
                if (pcd.encoding == Encoding.Unicode)
                    pcdFcEnd *= 2;
                pcdFcEnd += (Int32)pcd.fc;

                if (pcdFcEnd < fcStart)
                {
                    //this piece is before the requested range
                    continue;
                }
                else if (fcStart >= pcd.fc && fcEnd > pcdFcEnd)
                {
                    //requested char range starts at this piece
                    //read from fcStart to pcdFcEnd

                    //get count of bytes
                    int cb = pcdFcEnd - fcStart;
                    byte[] bytes = new byte[cb];

                    //read all bytes
                    wordStream.Read(bytes, 0, cb, (Int32)fcStart);

                    //get the chars
                    char[] plainChars = pcd.encoding.GetString(bytes).ToCharArray();

                    //add to list
                    foreach (char c in plainChars)
                    {
                        chars.Add(c);
                    }
                }
                else if (fcStart <= pcd.fc && fcEnd >= pcdFcEnd)
                {
                    //the full piece is part of the requested range
                    //read from pc.fc to pcdFcEnd

                    //get count of bytes
                    int cb = pcdFcEnd - (Int32)pcd.fc;
                    byte[] bytes = new byte[cb];

                    //read all bytes 
                    wordStream.Read(bytes, 0, cb, (Int32)pcd.fc);

                    //get the chars
                    char[] plainChars = pcd.encoding.GetString(bytes).ToCharArray();

                    //add to list
                    foreach (char c in plainChars)
                    {
                        chars.Add(c);
                    }
                }
                else if (fcStart < pcd.fc && fcEnd >= pcd.fc && fcEnd <= pcdFcEnd)
                {
                    //requested char range ends at this piece
                    //read from pcd.fc to fcEnd

                    //get count of bytes
                    int cb = fcEnd - (Int32)pcd.fc;
                    byte[] bytes = new byte[cb];

                    //read all bytes 
                    wordStream.Read(bytes, 0, cb, (Int32)pcd.fc);

                    //get the chars
                    char[] plainChars = pcd.encoding.GetString(bytes).ToCharArray();

                    //add to list
                    foreach (char c in plainChars)
                    {
                        chars.Add(c);
                    }

                    break;
                }
                else if (fcStart >= pcd.fc && fcEnd <= pcdFcEnd)
                {
                    //requested chars are completly in this piece
                    //read from fcStart to fcEnd

                    //get count of bytes
                    int cb = fcEnd - fcStart;
                    byte[] bytes = new byte[cb];

                    //read all bytes 
                    wordStream.Read(bytes, 0, cb, (Int32)fcStart);

                    //get the chars
                    char[] plainChars = pcd.encoding.GetString(bytes).ToCharArray();

                    //set the list
                    chars = new List<char>(plainChars);

                    break;
                }
                else if (fcEnd < pcd.fc)
                {
                    //this piece is beyond the requested range
                    break;
                }
            }
            return chars;
        }
    }
}
