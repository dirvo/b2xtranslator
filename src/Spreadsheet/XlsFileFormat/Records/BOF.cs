/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies properties about the substream and specifies the beginning 
    /// of a collection of records as defined by the Workbook Stream ABNF and the Chart Sheet Substream ABNF.
    /// </summary>
    [BiffRecordAttribute(RecordType.BOF)]
    public class BOF : BiffRecord
    {
        public const RecordType ID = RecordType.BOF;

        public enum DocumentType : ushort
        {
            /// <summary>
            /// Workbook globals
            /// </summary>
            WorkbookGlobals = 0x005,

            /// <summary>
            /// Visual Basic module
            /// </summary>
            VisualBasicModule = 0x0006,

            /// <summary>
            /// Worksheet or dialog sheet
            /// </summary>
            Worksheet = 0x0010,

            /// <summary>
            /// Chart
            /// </summary>
            Chart = 0x0020,

            /// <summary>
            /// Excel 4.0 macro sheet
            /// </summary>
            MacroSheet = 0x0040,

            /// <summary>
            /// Workspace file
            /// </summary>
            WorkspaceFile = 0x0100,

            /// <summary>
            /// from MS-OGRAPH spec
            /// </summary>
            ChartSheet = 0x8000
        }

        /// <summary>
        /// An unsigned integer that specifies the version of the substream. 
        /// 
        /// MUST be 0x0680.
        /// </summary>
        public UInt16 version;

        /// <summary>
        /// An unsigned integer that specifies the type of data contained in the substream. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value     Meaning
        ///     0x0005    Specifies a workbook stream.
        ///     0x8000    Specifies a chart sheet substream.
        /// </summary>
        public DocumentType docType;

        /// <summary>
        /// An unsigned integer that specifies the build identifier of the application creating the substream.
        /// </summary>
        public UInt16 rupBuild;

        /// <summary>
        /// An unsigned integer that specifies the version of the file format. 
        /// 
        /// This value MUST be 0x07CC or 0x07CD. 
        /// This value SHOULD be 0x07CD (1997).
        /// </summary>
        public UInt16 rupYear;

        /// <summary>
        /// A bit that specifies whether this substream was last edited on a Windows platform. 
        /// 
        /// MUST be 1.
        /// </summary>
        public bool fWin;

        /// <summary>
        /// A bit that specifies whether the substream was last edited on a RISC platform. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fRisc;

        /// <summary>
        /// A bit that specifies whether this substream was last edited by a beta version of the application. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fBeta;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited on a Windows platform. 
        /// 
        /// MUST be 1.
        /// </summary>
        public bool fWinAny;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited on a Macintosh platform. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fMacAny;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited by a beta version of the application. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fBetaAny;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited on a RISC platform. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fRiscAny;

        /// <summary>
        /// A bit that specifies whether this substream caused an out-of-memory failure.
        /// </summary>
        public bool fOOM;

        /// <summary>
        /// A bit that specifies whether this substream caused an out-of-memory failure while loading charting or graphics data.
        /// </summary>
        public bool fGlJmp;

        /// <summary>
        /// A bit that specifies that whether this substream has hit the 255 font limit, such that new Font records cannot be added to it.
        /// </summary>
        public bool fFontLimit;

        /// <summary>
        /// An unsigned integer (4 bits) that specifies the highest version of the application that has ever saved this substream.
        /// </summary>
        public Byte verXLHigh;

        /// <summary>
        /// An unsigned integer that specifies the version of the file format. 
        /// 
        /// MUST be 0x06.
        /// </summary>
        public Byte verLowestBiff;

        /// <summary>
        /// An unsigned integer (4 bits) that specifies the application version that saved this substream most recently. The value MUST be the value of verXLHigh field or less.
        /// </summary>
        public Byte verLastXLSaved;

        public BOF(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.version = reader.ReadUInt16();
            this.docType = (DocumentType)reader.ReadUInt16();
            this.rupBuild = reader.ReadUInt16();
            this.rupYear = reader.ReadUInt16();

            UInt32 flags = reader.ReadUInt32();
            this.fWin = Utils.BitmaskToBool(flags, 0x0001);
            this.fRisc = Utils.BitmaskToBool(flags, 0x0002);
            this.fBeta = Utils.BitmaskToBool(flags, 0x0004);
            this.fWinAny = Utils.BitmaskToBool(flags, 0x0008);
            this.fMacAny = Utils.BitmaskToBool(flags, 0x0010);
            this.fBetaAny = Utils.BitmaskToBool(flags, 0x0020);
            // 2 bits ignored
            this.fRiscAny = Utils.BitmaskToBool(flags, 0x0100);
            this.fOOM = Utils.BitmaskToBool(flags, 0x0200);
            this.fGlJmp = Utils.BitmaskToBool(flags, 0x0400);
            // 2 bits ignored
            this.fFontLimit = Utils.BitmaskToBool(flags, 0x2000);
            this.verXLHigh = Utils.BitmaskToByte(flags, 0x0003C000);

            this.verLowestBiff = reader.ReadByte();
            this.verLastXLSaved = Utils.BitmaskToByte(reader.ReadUInt16(), 0x00FF);

            // ignore remaing part of record
            reader.ReadByte();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
