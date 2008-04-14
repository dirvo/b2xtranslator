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
    public class ListLevel
    {
        public enum FollowingChar
        {
            tab = 0,
            space,
            nothing
        }

        /// <summary>
        /// Start at value for this list level
        /// </summary>
        public Int32 iStartAt;

        /// <summary>
        /// Number format code (see anld.nfc for a list of options)
        /// </summary>
        public byte nfc;

        /// <summary>
        /// Alignment (left, right, or centered) of the paragraph number.
        /// </summary>
        public byte jc;

        /// <summary>
        /// True if the level turns all inherited numbers to arabic, 
        /// false if it preserves their number format code (nfc)
        /// </summary>
        public bool fLegal;

        /// <summary>
        /// True if the level‘s number sequence is not restarted by 
        /// higher (more significant) levels in the list
        /// </summary>
        public bool fNoRestart;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.fPrev (see ANLD)
        /// </summary>
        public bool fPrev;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.fPrevSpace (see ANLD)
        /// </summary>
        public bool fPrevSpace;

        /// <summary>
        /// True if this level was from a converted Word 6.0 document. <br/>
        /// If it is true, all of the Word 6.0 compatibility options become 
        /// valid otherwise they are ignored.
        /// </summary>
        public bool fWord6;

        /// <summary>
        /// Contains the character offsets into the LVL’s XST of the inherited numbers of previous levels. <br/>
        /// The XST contains place holders for any paragraph numbers contained in the text of the number, 
        /// and the place holder contains the ilvl of the inherited number, 
        /// so lvl.xst[lvl.rgbxchNums[0]] == the level of the first inherited number in this level.
        /// </summary>
        public byte[] rgbxchNums;

        /// <summary>
        /// The type of character following the number text for the paragraph.
        /// </summary>
        public FollowingChar ixchFollow;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.dxaSpace (see ANLD). <br/>
        /// For newer versions indent to remove if we remove this numbering.
        /// </summary>
        public Int32 dxaSpace;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.dxaIndent (see ANLD).<br/>
        /// Unused in newer versions.
        /// </summary>
        public Int32 dxaIndent;

        /// <summary>
        /// Length, in bytes, of the LVL‘s grpprlChpx.
        /// </summary>
        public byte cbGrpprlChpx;

        /// <summary>
        /// Length, in bytes, of the LVL‘s grpprlPapx.
        /// </summary>
        public byte cbGrpprlPapx;

        /// <summary>
        /// Limit of levels that we restart after.
        /// </summary>
        public byte ilvlRestartLim;

        /// <summary>
        /// 
        /// </summary>
        public ParagraphPropertyExceptions grpprlPapx;

        /// <summary>
        /// 
        /// </summary>
        public CharacterPropertyExceptions grpprlChpx;

        /// <summary>
        /// 
        /// </summary>
        public string xst;

        private const int LVLF_LENGTH = 28;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bytes"></param>
        public ListLevel(VirtualStream tableStream, int offset)
        { 
            //read the fix part of the LVL into an array
            byte[] bytes = new byte[LVLF_LENGTH];
            tableStream.Read(bytes, 0, LVLF_LENGTH, offset);
            
            //parse the fix part
            this.iStartAt = System.BitConverter.ToInt32(bytes, 0);
            this.nfc = bytes[4];
            int flag = (int)bytes[5];
            this.jc = (byte)(flag & 0x03);
            this.fLegal = Utils.BitmaskToBool(flag, 0x04);
            this.fNoRestart = Utils.BitmaskToBool(flag, 0x08);
            this.fPrev = Utils.BitmaskToBool(flag, 0x10);
            this.fPrevSpace = Utils.BitmaskToBool(flag, 0x20);
            this.fWord6 = Utils.BitmaskToBool(flag, 0x40);
            this.rgbxchNums = new byte[9];
            int j=0;
            for (int i = 6; i < 15; i++)
            {
                rgbxchNums[j] = bytes[i];
                j++;
            }
            this.ixchFollow = (FollowingChar)bytes[15];
            this.dxaSpace = System.BitConverter.ToInt32(bytes, 16);
            this.dxaIndent = System.BitConverter.ToInt32(bytes, 20);
            this.cbGrpprlChpx = bytes[24];
            this.cbGrpprlPapx = bytes[25];
            this.ilvlRestartLim = bytes[26];

            //parse the variable part

            //read the group of papx sprms
            int papxPos = offset + LVLF_LENGTH;
            byte[] papxBytes = new byte[this.cbGrpprlPapx];
            tableStream.Read(papxBytes, 0, this.cbGrpprlPapx, papxPos);
            //this papx ahs no istd, so use PX to parse it
            PropertyExceptions px = new PropertyExceptions(papxBytes);
            this.grpprlPapx = new ParagraphPropertyExceptions();
            this.grpprlPapx.grpprl = px.grpprl;

            //read the group of chpx sprms
            int chpxPos = offset + LVLF_LENGTH + this.cbGrpprlPapx;
            byte[] chpxBytes = new byte[this.cbGrpprlChpx];
            tableStream.Read(chpxBytes, 0, this.cbGrpprlChpx, chpxPos);
            this.grpprlChpx = new CharacterPropertyExceptions(chpxBytes);

            //read the number text
            int strPos = offset + LVLF_LENGTH + this.cbGrpprlPapx + this.cbGrpprlChpx;
            byte[] strLenBytes = new byte[2];
            tableStream.Read(strLenBytes,0,2,strPos);
            Int16 strLen = System.BitConverter.ToInt16(strLenBytes, 0);
            strPos += 2;
            byte[] strBytes = new byte[strLen*2];
            tableStream.Read(strBytes, 0, strLen*2, strPos);
            this.xst = Encoding.Unicode.GetString(strBytes);
        }
    }
}
