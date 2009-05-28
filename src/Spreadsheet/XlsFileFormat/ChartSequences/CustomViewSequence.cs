using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class CustomViewSequence : BiffRecordSequence
    {
        public CustomViewSequence(IStreamReader reader)
            : base(reader)
        {

        }
    }
}
