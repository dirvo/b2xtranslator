using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    public class ChartFormats : OfficeGraphBiffRecordSequence
    {
        public Chart Chart;

        public ChartFormats(IStreamReader reader) : base(reader)
        {

        }
    }
}
