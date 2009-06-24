﻿using System;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.SIIndex)]
    public class SIIndex : BiffRecord
    {
        public enum SeriesDataType : ushort
        {
            SeriesValues = 0x0001,
            CategoryLabels = 0x0002,
            BubbleSizes = 0x0003
        }

        /// <summary>
        /// An unsigned integer that specifies the type of the data records contained by the Number records following it. <br/>
        /// MUST be a value from the following table:
        /// </summary>
        public SeriesDataType numIndex;

        public SIIndex(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            this.numIndex = (SeriesDataType)reader.ReadUInt16();
        }
    }
}
