﻿using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class BackgroundSequence : BiffRecordSequence
    {
        public BkHim BkHim;

        public List<Continue> Continues;

        public BackgroundSequence(IStreamReader reader)
            : base(reader)
        {
            // BACKGROUND = BkHim *Continue

            // BkHim
            this.BkHim = (BkHim)BiffRecord.ReadRecord(reader);

            // *Continue
            this.Continues = new List<Continue>();
            while(BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continues.Add((Continue)BiffRecord.ReadRecord(reader));
            }

        }
    }
}
