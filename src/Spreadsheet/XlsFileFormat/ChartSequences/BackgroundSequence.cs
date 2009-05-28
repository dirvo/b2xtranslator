using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class BackgroundSequence : BiffRecordSequence
    {
        public BackgroundSequence(IStreamReader reader)
            : base(reader)
        {
            // BACKGROUND = BkHim *Continue
        }
    }
}
