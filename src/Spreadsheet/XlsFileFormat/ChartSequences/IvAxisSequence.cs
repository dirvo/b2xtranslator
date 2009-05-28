using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class IvAxisSequence : BiffRecordSequence
    {
        public Axis Axis;

        public Begin Begin;

        public CatSerRange CatSerRange;

        public AxcExt AxcExt;

        public CatLab CatLab;

        public AxsSequence AxsSequence;

        public FutureRecordSequence FutureRecordSequence;

        public End End;

        public IvAxisSequence(IStreamReader reader)
            : base(reader)
        {
            // IVAXIS = Axis Begin [CatSerRange] AxcExt [CatLab] AXS [CRTMLFRT] End

            // Axis
            this.Axis = (Axis)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // [CatSerRange]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CatSerRange)
            {
                this.CatSerRange = (CatSerRange)BiffRecord.ReadRecord(reader);
            }

            // AxcExt
            this.AxcExt = (AxcExt)BiffRecord.ReadRecord(reader);

            // [CatLab]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CatLab)
            {
                this.CatLab = (CatLab)BiffRecord.ReadRecord(reader);
            }

            // AXS
            this.AxsSequence = new AxsSequence(reader);

            // [CRTMLFRT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.FutureRecordSequence = new FutureRecordSequence(reader);
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
