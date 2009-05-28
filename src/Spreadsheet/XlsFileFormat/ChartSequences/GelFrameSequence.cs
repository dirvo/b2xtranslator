using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class GelFrameSequence : BiffRecordSequence
    {
        public List<GelFrame> GelFrames;

        public List<Continue> Continues;

        public PicfSequence PicfSequence;

        public GelFrameSequence(IStreamReader reader)
            : base(reader)
        {
            // GELFRAME = 1*2GelFrame *Continue [PICF]

            // 1*2GelFrame
            this.GelFrames = new List<GelFrame>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
            {
                this.GelFrames.Add((GelFrame)BiffRecord.ReadRecord(reader));
            }

            // *Continue
            this.Continues = new List<Continue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continues.Add((Continue)BiffRecord.ReadRecord(reader));
            }

            if (BiffRecord.GetNextRecordType(reader) == RecordType.Begin)
            {
                this.PicfSequence = new PicfSequence(reader);
            }
        }
    }
}
