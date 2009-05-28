﻿using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AiSequence : BiffRecordSequence
    {
        public BRAI BRAI;

        public SeriesText SeriesText;

        public AiSequence(IStreamReader reader)
            : base(reader)
        {
            //AI = BRAI [SeriesText]

            //BRAI
            this.BRAI = (BRAI)BiffRecord.ReadRecord(reader);

            //[SeriesText]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.SeriesText)
            {
                this.SeriesText = (SeriesText)BiffRecord.ReadRecord(reader);
            }

        }
    }
}
