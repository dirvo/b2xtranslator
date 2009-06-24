﻿using System;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.SBaseRef)]
    public class SBaseRef : BiffRecord
    {
        /// <summary>
        ///  A RwU that specifies the zero-based index of the first row in the range. <br/>
        ///  MUST be less than or equal to rwLast.
        /// </summary>
        public UInt16 rwFirst;

        /// <summary>
        /// A RwU that specifies the zero-based index of the last row in the range. <br/>
        /// MUST be greater than or equal to rwFirst.
        /// </summary>
        public UInt16 rwLast;

        /// <summary>
        /// A ColU that specifies the zero-based index of the first column in the range.<br/> 
        /// MUST be less than or equal to colLast.<br/>
        /// MUST be less than or equal to 0x00FF.
        /// </summary>
        public UInt16 colFirst;

        /// <summary>
        /// A ColU that specifies the zero-based index of the last column in the range. <br/>
        /// MUST be greater than or equal to colFirst.<br/>
        /// MUST be less than or equal to 0x00FF.
        /// </summary>
        public UInt16 colLast;

        public SBaseRef(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();
        }
    }
}
