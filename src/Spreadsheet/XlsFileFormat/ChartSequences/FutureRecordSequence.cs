using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class FutureRecordSequence : OfficeGraphBiffRecordSequence
    {
        public FutureRecordSequence(IStreamReader reader)
            : base(reader)
        {
        }
    }
}
