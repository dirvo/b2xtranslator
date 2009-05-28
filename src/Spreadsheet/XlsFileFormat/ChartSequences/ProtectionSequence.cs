﻿using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ProtectionSequence : BiffRecordSequence
    {
        public Protect Protect;

        public ScenarioProtect ScenarioProtect;

        public ObjProtect ObjProtect;

        public Password Password;

        public ProtectionSequence(IStreamReader reader)
            : base(reader)
        {
            // PROTECTION = [Protect] [ScenarioProtect] [ObjProtect] [Password]

            // [Protect]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Protect)
            {
                this.Protect = (Protect)BiffRecord.ReadRecord(reader);
            }

            // [ScenarioProtect]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ScenarioProtect)
            {
                this.ScenarioProtect = (ScenarioProtect)BiffRecord.ReadRecord(reader);
            }

            // [ObjProtect]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ObjProtect)
            {
                this.ObjProtect = (ObjProtect)BiffRecord.ReadRecord(reader);
            }

            // [Password]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Password)
            {
                this.Password = (Password)BiffRecord.ReadRecord(reader);
            }
        }
    }
}
