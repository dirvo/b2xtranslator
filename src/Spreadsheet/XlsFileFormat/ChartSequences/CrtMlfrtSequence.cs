using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class CrtMlfrtSequence : BiffRecordSequence
    {
        public CrtMlFrt CrtMlFrt;

        public List<CrtMlFrtContinue> CrtMlFrtContinues;

        public CrtMlfrtSequence(IStreamReader reader)
            : base(reader)
        {
            //CRTMLFRT = CrtMlFrt *CrtMlFrtContinue

            this.CrtMlFrt = (CrtMlFrt)BiffRecord.ReadRecord(reader);

            this.CrtMlFrtContinues = new List<CrtMlFrtContinue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrtContinue)
            {
                this.CrtMlFrtContinues.Add((CrtMlFrtContinue)BiffRecord.ReadRecord(reader));
            }

        }
    }
}
