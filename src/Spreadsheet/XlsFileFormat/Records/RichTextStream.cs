using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.RichTextStream)]
    public class RichTextStream : BiffRecord
    {
        public RichTextStream(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {

        }
    }
}
