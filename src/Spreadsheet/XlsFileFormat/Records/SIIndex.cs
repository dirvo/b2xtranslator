using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.SIIndex)]
    public class SIIndex : BiffRecord
    {
        /// <summary>
        /// An unsigned integer that specifies the type of the data records contained by the Number records following it. <br/>
        /// MUST be a value from the following table:
        /// </summary>
        public UInt16 numIndex;

        public SIIndex(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            this.numIndex = reader.ReadUInt16();
        }
    }
}
