using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using OfficeGraph = DIaLOGIKa.b2xtranslator.OfficeGraph;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartFormatsSequence : OfficeGraphBiffRecordSequence
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

        public List<FutureRecordSequence> FutureRecordSequences;

        public End End;

        public ChartFormatsSequence(IStreamReader reader)
            : base(reader)
        {
           // CHARTFOMATS = Chart Begin *2FONTLIST Scl PlotGrowth [FRAME] *SERIESFORMAT *SS ShtProps 
           //     *2DFTTEXT AxesUsed 1*2AXISPARENT [CrtLayout12A] [DAT] *ATTACHEDLABEL [CRTMLFRT]
           //     *([DataLabExt StartObject] ATTACHEDLABEL [EndObject]) [TEXTPROPS] *2CRTMLFRT End


            // Chart
            this.Chart = (Chart)OfficeGraphBiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // *2FONTLIST
            this.FontListSequences = new List<FontListSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.FrtFontList)
            {
                this.FontListSequences.Add(new FontListSequence(reader));
            }

            // Scl
            this.Scl = (Scl)OfficeGraphBiffRecord.ReadRecord(reader);

            // PlotGrowth
            this.PlotGrowth = (PlotGrowth)OfficeGraphBiffRecord.ReadRecord(reader);

            // [FRAME]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.Frame)
            {
                this.FrameSequence = new FrameSequence(reader);
            }

            // *SERIESFORMAT
            this.SeriesFormatSequences = new List<SeriesFormatSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.Series)
            {
                this.SeriesFormatSequences.Add(new SeriesFormatSequence(reader));
            }

            // *SS
            this.SsSequences = new List<SsSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.DataFormat)
            {
                this.SsSequences.Add(new SsSequence(reader));
            }

            // ShtProps
            this.ShtProps = (ShtProps)OfficeGraphBiffRecord.ReadRecord(reader);

            // *2DFTTEXT
            this.DftTextSequences = new List<DftTextSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.DataLabExt)
            {
                this.DftTextSequences.Add(new DftTextSequence(reader));
            }

            // AxesUsed
            this.AxesUsed = (AxesUsed)OfficeGraphBiffRecord.ReadRecord(reader);

            // 1*2AXISPARENT
            this.AxisParentSequences = new List<AxisParentSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.AxisParent)
            {
                this.AxisParentSequences.Add(new AxisParentSequence(reader));
            }

            // [CrtLayout12A]

            // [DAT]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.Dat)
            {
                this.Dat = (Dat)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // *ATTACHEDLABEL
            this.AttachedLabels = new List<AttachedLabel>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.AttachedLabel)
            {
                this.AttachedLabels.Add((AttachedLabel)OfficeGraphBiffRecord.ReadRecord(reader));
            }

            // [CRTMLFRT]

            // *([DataLabExt StartObject] ATTACHEDLABEL [EndObject])
            this.DataLabelGroups = new List<DataLabelGroup>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.DataLabExt)
            {
               this.DataLabelGroups.Add(new DataLabelGroup(reader));
            }

            // [TEXTPROPS]
            //if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) )
            //{
            //    this.TextPropsSequence = new TextPropsSequence(reader);
            //}

            // *2CRTMLFRT
            this.FutureRecordSequences = new List<FutureRecordSequence>();


            // End
            this.End = (End)OfficeGraphBiffRecord.ReadRecord(reader);
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

                if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.DataLabExt)
                {
                    this.DataLabExt = (DataLabExt)OfficeGraphBiffRecord.ReadRecord(reader);
                    this.StartObject = (StartObject)OfficeGraphBiffRecord.ReadRecord(reader);
                }

                this.AttachedLabel = (AttachedLabel)OfficeGraphBiffRecord.ReadRecord(reader);

                if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.EndObject)
                {
                    this.EndObject = (EndObject)OfficeGraphBiffRecord.ReadRecord(reader);
                }
            }
        }
    }

}
