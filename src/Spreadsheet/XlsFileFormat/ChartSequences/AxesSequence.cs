using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AxesSequence : BiffRecordSequence
    {
        public IvAxisSequence IvAxisSequence;

        public DvAxisSequence DvAxisSequence;

        public DvAxisSequence DvAxisSequence2;

        public SeriesAxisSequence SeriesAxisSequence;

        public List<AttachedLabelSequence> AttachedLabelSequences;

        public PlotArea PlotArea;

        public FrameSequence Frame;

        public AxesSequence(IStreamReader reader)
            : base(reader)
        {
            //AXES = [IVAXIS DVAXIS [SERIESAXIS] / DVAXIS DVAXIS] *3ATTACHEDLABEL [PlotArea FRAME]

            // [IVAXIS DVAXIS [SERIESAXIS] / DVAXIS DVAXIS]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Axis)
            {
                long position = reader.BaseStream.Position;

                Axis axis = (Axis)BiffRecord.ReadRecord(reader);

                Begin begin = (Begin)BiffRecord.ReadRecord(reader);

                if (BiffRecord.GetNextRecordType(reader) == RecordType.CatSerRange
                    || BiffRecord.GetNextRecordType(reader) == RecordType.AxcExt)
                {
                    reader.BaseStream.Position = position;

                    //IVAXIS
                    this.IvAxisSequence = new IvAxisSequence(reader);

                    //DVAXIS 
                    this.DvAxisSequence = new DvAxisSequence(reader);
                    
                    //[SERIESAXIS]  
                    if (BiffRecord.GetNextRecordType(reader) == RecordType.Axis)
                    {
                        this.SeriesAxisSequence = new SeriesAxisSequence(reader);
                    }
                }
                else
                {
                    reader.BaseStream.Position = position;

                    //DVAXIS 
                    this.DvAxisSequence = new DvAxisSequence(reader);
                    
                    //DVAXIS 
                    this.DvAxisSequence2 = new DvAxisSequence(reader);
                }
            }

            //*3ATTACHEDLABEL 
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Text)
            {
                this.AttachedLabelSequences.Add(new AttachedLabelSequence(reader));
            }

            //[PlotArea FRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PlotArea)
            {
                this.PlotArea = (PlotArea)BiffRecord.ReadRecord(reader);

                this.Frame = new FrameSequence(reader);
            }
        }
    }
}
