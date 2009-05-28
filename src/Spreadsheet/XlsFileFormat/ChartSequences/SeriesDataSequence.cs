using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class SeriesDataSequence : BiffRecordSequence
    {
        public Dimensions Dimensions;

        public SeriesGroup[] SeriesGroups;

        public SeriesDataSequence(IStreamReader reader)
            : base(reader)
        {
            // SERIESDATA = Dimensions 3(SIIndex *(Number / BoolErr / Blank / Label))

            // Dimensions
            this.Dimensions = (Dimensions)BiffRecord.ReadRecord(reader);

            // 3(SIIndex *(Number / BoolErr / Blank / Label))
            this.SeriesGroups = new SeriesGroup[3];
            for (int i = 0; i < 3; i++)
            {
                this.SeriesGroups[i] = new SeriesGroup(reader);
            }
        }
    }

    public class SeriesGroup
    {
        public SIIndex SIIndex;

        public List<BiffRecord> Data;

        public SeriesGroup(IStreamReader reader)
        {
            // SIIndex *(Number / BoolErr / Blank / Label)

            // SIIndex
            this.SIIndex = (SIIndex)BiffRecord.ReadRecord(reader);

            // *(Number / BoolErr / Blank / Label)
            this.Data = new List<BiffRecord>();
            while (
                BiffRecord.GetNextRecordType(reader) == RecordType.Number ||
                BiffRecord.GetNextRecordType(reader) == RecordType.BoolErr ||
                BiffRecord.GetNextRecordType(reader) == RecordType.Blank ||
                BiffRecord.GetNextRecordType(reader) == RecordType.Label)
            {
                this.Data.Add(BiffRecord.ReadRecord(reader));
            }
        }
    }
}
