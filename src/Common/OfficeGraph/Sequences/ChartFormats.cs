using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph.Sequences
{
    public class ChartFormats : RecordSequence
    {
        public Chart Chart;

        public ChartFormats(IStreamReader reader) : base(reader)
        {
        }
    }
}
