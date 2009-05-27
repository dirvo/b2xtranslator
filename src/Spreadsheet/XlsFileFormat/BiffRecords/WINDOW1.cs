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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Diagnostics;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords
{
    /// <summary>
    /// WINDOW1: Window Information (3Dh)
    /// 
    /// The WINDOW1 record contains workbook-level window attributes. 
    /// 
    /// The xWn and yWn fields contain the location of the window in units of 1/20th of a point, 
    /// relative to the upper-left corner of the Excel window client area. 
    /// 
    /// The dxWn and dyWn fields contain the window size, also in units of 1/20th  of a point.
    /// </summary>
    [BiffRecordAttribute(RecordNumber.WINDOW1)] 
    public class WINDOW1 : BiffRecord
    {
        public const RecordNumber ID = RecordNumber.WINDOW1;

        /// <summary>
        /// Horizontal position of the window in units of 1/20th of a point.
        /// </summary>
        public UInt16 xWn;
	
        /// <summary>
        /// Vertical position of the window in units of 1/20th of a point.
        /// </summary>
        public UInt16 yWn;

        /// <summary>
        /// Width of the window in units of 1/20th of a point.
        /// </summary>
        public UInt16 dxWn;
	
        /// <summary>
        /// Height of the window in units of 1/20th of a point.
        /// </summary>
        public UInt16 dyWn;
	
        /// <summary>
        /// Option flags
        /// </summary>
        public UInt16 grbit;
	
    	/// <summary>
    	/// Index of the selected workbook tab (0-based).
    	/// </summary>
        public UInt16 itabCur;
	
        /// <summary>
        /// Index of the first displayed workbook tab (0-based).
        /// </summary>
        public UInt16 itabFirst;
	
        /// <summary>
        /// Number of workbook tabs that are selected.
        /// </summary>
        public UInt16 ctabSel;

        /// <summary>
        /// Ratio of the width of the workbook tabs to the width of the horizontal scroll bar; 
        /// to obtain the ratio, convert to decimal and then divide by 1000.
        /// </summary>
        public UInt16 wTabRatio;	

        // The grbit field contains the following option flags:
        // Field                        Offset	Bits    Mask	Name	Contents
        public bool fHidden;        //  0	    0       01h		=1 if the window is hidden
        public bool fIconic; 	    //     	    1	    02h     =1 if the window is currently displayed as an icon
        public bool reserved0;      //     	    2	    04h	
        public bool fDspHScroll; 	//     	    3	    08h	    =1 if the horizontal scroll bar is displayed
        public bool fDspVScroll;    //    	    4	    10h	    =1 if the vertical scroll bar is displayed
        public bool fBotAdornment;  //     	    5	    20h	    =1 if the workbook tabs are displayed
        public bool fNoAFDateGroup; //     	    6	    40h	    =1 if the AutoFilter should not group dates (Excel 11 (2003) behavior), new for Office Excel 2007
        public bool reserved1;      //     	    7	    80h
        public byte reserved2;      //  1       7�0	    FFh	
        
        public WINDOW1(IStreamReader reader, RecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            xWn = reader.ReadUInt16();
            yWn = reader.ReadUInt16();
            dxWn = reader.ReadUInt16();
            dyWn = reader.ReadUInt16();

            grbit = reader.ReadUInt16();

            fHidden = Utils.BitmaskToBool(grbit, 0x01);
            fIconic = Utils.BitmaskToBool(grbit, 0x02);
            reserved0 = Utils.BitmaskToBool(grbit, 0x04);
            fDspHScroll = Utils.BitmaskToBool(grbit, 0x08);
            fDspVScroll = Utils.BitmaskToBool(grbit, 0x10);
            fBotAdornment = Utils.BitmaskToBool(grbit, 0x20);
            fNoAFDateGroup = Utils.BitmaskToBool(grbit, 0x40);
            reserved1 = Utils.BitmaskToBool(grbit, 0x80);
            reserved2 = Utils.BitmaskToByte(grbit, 0xFF00);

            itabCur = reader.ReadUInt16();
            itabFirst = reader.ReadUInt16();
            ctabSel = reader.ReadUInt16();
            wTabRatio = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
