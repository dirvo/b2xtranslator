using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class TextPropsSequence : BiffRecordSequence
    {
        public TextPropsSequence(IStreamReader reader)
            : base(reader)
        {
            // TEXTPROPS = (RichTextStream / TextPropsStream) *ContinueFrt12
        }
    }
}
