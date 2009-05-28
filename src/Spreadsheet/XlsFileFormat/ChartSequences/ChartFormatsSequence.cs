using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartFormatsSequence : BiffRecordSequence
    {
        public Chart Chart;

        public Begin Begin;

        public List<FontListSequence> FontListSequences;

        public Scl Scl;

        public PlotGrowth PlotGrowth;

        public FrameSequence FrameSequence;

        public List<SeriesFormatSequence> SeriesFormatSequences;

        public List<SsSequence> SsSequences;

        public ShtProps ShtProps;

        public List<DftTextSequence> DftTextSequences;

        public AxesUsed AxesUsed;

        public List<AxisParentSequence> AxisParentSequences;

        public List<AttachedLabel> AttachedLabels;

        public List<DataLabelGroup> DataLabelGroups;

        public TextPropsSequence TextPropsSequence;

        public Dat Dat;

        public CrtLayout12 CrtLayout12A;

        public CrtMlfrtSequence CrtMlfrtSequence;

        public List<CrtMlfrtSequence> CrtMlfrtSequences;

        public End End;

        public ChartFormatsSequence(IStreamReader reader)
            : base(reader)
        {
            // CHARTFOMATS = Chart Begin *2FONTLIST Scl PlotGrowth [FRAME] *SERIESFORMAT *SS ShtProps 
            //    *2DFTTEXT AxesUsed 1*2AXISPARENT [CrtLayout12A] [DAT] *ATTACHEDLABEL [CRTMLFRT] 
            //    *([DataLabExt StartObject] ATTACHEDLABEL [EndObject]) [TEXTPROPS] *2CRTMLFRT End

            // Chart
            this.Chart = (Chart)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // *2FONTLIST
            this.FontListSequences = new List<FontListSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.FrtFontList)
            {
                this.FontListSequences.Add(new FontListSequence(reader));
            }

            // Scl
            this.Scl = (Scl)BiffRecord.ReadRecord(reader);

            // PlotGrowth
            this.PlotGrowth = (PlotGrowth)BiffRecord.ReadRecord(reader);

            // [FRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Frame)
            {
                this.FrameSequence = new FrameSequence(reader);
            }

            // *SERIESFORMAT
            this.SeriesFormatSequences = new List<SeriesFormatSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Series)
            {
                this.SeriesFormatSequences.Add(new SeriesFormatSequence(reader));
            }

            // *SS
            this.SsSequences = new List<SsSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataFormat)
            {
                this.SsSequences.Add(new SsSequence(reader));
            }

            // ShtProps
            this.ShtProps = (ShtProps)BiffRecord.ReadRecord(reader);

            // *2DFTTEXT
            this.DftTextSequences = new List<DftTextSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt
                || BiffRecord.GetNextRecordType(reader) == RecordType.DefaultText)
            {
                this.DftTextSequences.Add(new DftTextSequence(reader));
            }

            // AxesUsed
            this.AxesUsed = (AxesUsed)BiffRecord.ReadRecord(reader);

            // 1*2AXISPARENT
            this.AxisParentSequences = new List<AxisParentSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.AxisParent)
            {
                this.AxisParentSequences.Add(new AxisParentSequence(reader));
            }

            // [CrtLayout12A]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12A)
            {
                this.CrtLayout12A = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }

            // [DAT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Dat)
            {
                this.Dat = (Dat)BiffRecord.ReadRecord(reader);
            }

            // *ATTACHEDLABEL
            this.AttachedLabels = new List<AttachedLabel>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.AttachedLabel)
            {
                this.AttachedLabels.Add((AttachedLabel)BiffRecord.ReadRecord(reader));
            }

            // [CRTMLFRT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }

            // *([DataLabExt StartObject] ATTACHEDLABEL [EndObject])
            this.DataLabelGroups = new List<DataLabelGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt)
            {
               this.DataLabelGroups.Add(new DataLabelGroup(reader));
            }

            // [TEXTPROPS]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.RichTextStream)
            {
                this.TextPropsSequence = new TextPropsSequence(reader);
            }

            // *2CRTMLFRT
            this.CrtMlfrtSequences = new List<CrtMlfrtSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequences.Add(new CrtMlfrtSequence(reader));
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }

        public class DataLabelGroup
        {
            public DataLabExt DataLabExt;
            public StartObject StartObject;
            public AttachedLabel AttachedLabel;
            public EndObject EndObject;

            public DataLabelGroup(IStreamReader reader)
            {
                // *([DataLabExt StartObject] ATTACHEDLABEL [EndObject])

                if (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt)
                {
                    this.DataLabExt = (DataLabExt)BiffRecord.ReadRecord(reader);
                    this.StartObject = (StartObject)BiffRecord.ReadRecord(reader);
                }

                this.AttachedLabel = (AttachedLabel)BiffRecord.ReadRecord(reader);

                if (BiffRecord.GetNextRecordType(reader) == RecordType.EndObject)
                {
                    this.EndObject = (EndObject)BiffRecord.ReadRecord(reader);
                }
            }
        }
    }

}
