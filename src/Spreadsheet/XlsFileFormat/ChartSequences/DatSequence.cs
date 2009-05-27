using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class DatSequence : OfficeGraphBiffRecordSequence
    {
        public DatSequence(IStreamReader reader)
            : base(reader)
        {

        }
    }
}
