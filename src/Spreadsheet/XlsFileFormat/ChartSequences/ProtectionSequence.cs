using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ProtectionSequence : BiffRecordSequence
    {
        public ProtectionSequence(IStreamReader reader)
            : base(reader)
        {
            // PROTECTION = [Protect] [ScenarioProtect] [ObjProtect] [Password]
        }
    }
}
