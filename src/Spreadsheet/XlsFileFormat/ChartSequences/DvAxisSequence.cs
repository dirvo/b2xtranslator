using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class DvAxisSequence : BiffRecordSequence
    {
        public DvAxisSequence(IStreamReader reader)
            : base(reader)
        {

        }
    }
}
