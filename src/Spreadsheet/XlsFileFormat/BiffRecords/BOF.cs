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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// BOF: Beginning of File (809h)
    /// 
    /// The BOF record marks the beginning of the Book stream in the BIFF file. 
    /// It also marks the beginning of record groups (or “substreams” of the Book stream) 
    /// for sheets in the workbook. 
    /// 
    /// For BIFF2 through BIFF4, the BIFF version is found from the high-order byte of the 
    /// record number field, as shown in the following table. For BIFF5/BIFF7, 
    /// and BIFF8 use the vers field at offset 4 to determine the BIFF version.
    /// </summary>
    public class BOF : BiffRecord
    {
        public const RecordNumber ID = RecordNumber.BOF; 
        
        public UInt16 vers;         // Version number:
                                    //    0600 for BIFF8
        public UInt16 dt;	        // Substream type:
                                    //    0005h = Workbook globals
                                    //    0006h = Visual Basic module
                                    //    0010h = Worksheet or dialog sheet
                                    //    0020h = Chart
                                    //    0040h = Excel 4.0 macro sheet
                                    //    0100h = Workspace file
        public UInt16 rupBuild;     // Build identifier (=0DBBh for Excel 97)
        public UInt16 rupYear; 	    // Build year (=07CCh for Excel 97)
        public UInt32 bfh;          // File history flags
        public UInt32 sfo;          // Lowest BIFF version (see text)

        // The bfh field contains the following flag bits:
        // Field                    Bit     Mask
        public bool fWin;	    //  0	    00000001h   =1 if the file was last edited by Excel for Windows
        public bool fRisc;	    //  1	    00000002h   =1 if the file was last edited by Excel on a RISC platform
        public bool fBeta;	    //  2	    00000004h   =1 if the file was last edited by a beta version of Excel
        public bool fWinAny;    //  3	    00000008h   =1 if the file has ever been edited by Excel for Windows
        public bool fMacAny;    //  4	    00000010h   =1 if the file has ever been edited by Excel for the Macintosh
        public bool fBetaAny;   //  5	    00000020h   =1 if the file has ever been edited by a beta version of Excel
        public UInt32 reserved0;//  7–6 	000000C0h	Reserved; must be 0 (zero)
        public bool fRiscAny;   //  8	    00000100h   =1 if the file has ever been edited by Excel on a RISC platform
        public UInt32 reserved1;//  31-9    FFFFFE00h   Reserved; must be 0 (zero)


        public BOF(IStreamReader reader, RecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            vers = reader.ReadUInt16();

            // TODO: currently only BIFF8 is supported
            if (vers != 0x0600)
            {
                throw new NotSupportedException("Only BIFF8 files are supported.");
            }
            
            dt = reader.ReadUInt16();
            rupBuild = reader.ReadUInt16();
            rupYear = reader.ReadUInt16();
            bfh = reader.ReadUInt32();
            sfo = reader.ReadUInt32();

            fWin = Utils.BitmaskToBool(bfh, 0x00000001);
            fRisc = Utils.BitmaskToBool(bfh, 0x00000002);
            fBeta = Utils.BitmaskToBool(bfh, 0x00000004);
            fWinAny = Utils.BitmaskToBool(bfh, 0x00000008);
            fMacAny = Utils.BitmaskToBool(bfh, 0x00000010);
            fBetaAny = Utils.BitmaskToBool(bfh, 0x00000020);
            reserved0 = (uint)Utils.BitmaskToInt((int)bfh, 0x000000C0);
            fRiscAny = Utils.BitmaskToBool(bfh, 0x00000100);
            // reserved1 = (uint)Utils.BitmaskToInt((int)bfh, 0xFFFFFE00);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
