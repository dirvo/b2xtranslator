/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords
{
    public enum SuperSubScriptStyle : ushort
    {
        None, 
        Superscript, 
        Subscript 
    }

    public enum UnderlineStyle : ushort
    {
        None =              0x00,
        SingleLine =        0x01,
        DoubleLine =        0x02,
        SingleAccounting =  0x21,
        DoubleAccounting =  0x22
    }
    
    /// <summary>
    /// FONT: Font Description (231h)
    /// 
    /// The workbook font table contains at least five FONT records. 
    /// FONT records are numbered as follows: 
    ///     ifnt=00h (the first FONT record in the table), 
    ///     ifnt=01h, ifnt=02h, ifnt=03h, ifnt=05h (minimum table), 
    /// and then ifnt=06h, ifnt=07h, and so on.  
    /// 
    /// Note: ifnt=04h never appears in a BIFF file.  
    ///     This is for backward-compatibility with previous versions of Excel. 
    ///     If you read FONT records, remember to index the table correctly, skipping ifnt=04h.
    /// </summary>
    public class FONT : BiffRecord
    {
        public const RecordNumber ID = RecordNumber.FONT;

        /// <summary>
        /// Height of the font (in units of 1/20th of a point). 
        /// </summary>
        public UInt16 dyHeight;	

        /// <summary>
        /// Font attributes (packed bit field).
        /// </summary>
        public UInt16 grbit;	
    
        /// <summary>
        /// Index to the color palette.
        /// </summary>
        public UInt16 icv;	    

        /// <summary>
        /// Bold style; a number from 100dec to 1000dec (64h to 3E8h) 
        /// that indicates the character weight (“boldness”). 
        /// 
        /// The default values are 190h for normal text and 2BCh for bold text.
        /// </summary>
        public UInt16 bls;	    

        /// <summary>
        /// Superscript/subscript:
        ///     00h= None
        ///     01h= Superscript
        ///     02h= Subscript 
        /// </summary>
        public SuperSubScriptStyle sss;	    

        /// <summary>
        /// Underline style:
        ///     00h= None
        ///     01h= Single
        ///     02h= Double
        ///     21h= Single Accounting
        ///     22h= Double Accounting 
        /// </summary>
        public UnderlineStyle uls;	
     
        /// <summary>
        /// Font family, as defined by the Windows API LOGFONT structure.
        /// </summary>
        public byte bFamily;	
    
        /// <summary>
        /// Character set, as defined by the Windows API LOGFONT structure.
        /// </summary>
        public byte bCharSet;	

        /// <summary>
        /// Reserved; must be 0 (zero).
        /// </summary>
        public byte reserved0;	

        /// <summary>
        /// Length of the font name.
        /// </summary>
        public byte cch;	        

        /// <summary>
        /// Font name.
        /// </summary>
        public byte[] rgch;	   
 	
        // The grbit field contains the following font attributes:
        // Offset	Bits	Mask	Flag Name	Contents
        public bool fReserved0;  //  0	0	    01h	    Reserved; must be 0 (zero)
  	    public bool fItalic;	//      1	    02h	    =1 if the font is italic
  	    public bool fReserved1;	//      2	    04h	    Reserved; must be 0 (zero)
        public bool fStrikeout;	//  0	3	    08h	    =1 if the font is struck out
  	    public bool fOutline;	//      4	    10h	    =1 if the font is outline style (Macintosh only)
  	    public bool fShadow;	//      5	    20h	    =1 if the font is shadow style (Macintosh only)
  	    public byte fReserved2;	//      7 – 6 	C0h	    Reserved; must be 0 (zero)
        public byte fUnused0;	//  1	7 – 0 	FFh	    

        public FONT(IStreamReader reader, RecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            dyHeight = reader.ReadUInt16();
            grbit = reader.ReadUInt16();

            fReserved0 = Utils.BitmaskToBool(grbit, 0x01);
            fItalic = Utils.BitmaskToBool(grbit, 0x02);	  
            fReserved1 = Utils.BitmaskToBool(grbit, 0x04);
            fStrikeout = Utils.BitmaskToBool(grbit, 0x08);
            fOutline = Utils.BitmaskToBool(grbit, 0x10);	
            fShadow	= Utils.BitmaskToBool(grbit, 0x20);
            fReserved2 = Utils.BitmaskToByte(grbit, 0xC0);
            fUnused0	= Utils.BitmaskToByte(grbit, 0xFF00);  

            icv = reader.ReadUInt16();
            bls = reader.ReadUInt16();
            sss = (SuperSubScriptStyle)reader.ReadUInt16();
            uls = (UnderlineStyle)reader.ReadByte();
            bFamily = reader.ReadByte();
            bCharSet = reader.ReadByte();
            reserved0 = reader.ReadByte();
            cch = reader.ReadByte();
            rgch = reader.ReadBytes(cch);	 
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
