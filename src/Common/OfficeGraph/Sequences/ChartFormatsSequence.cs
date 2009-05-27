using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
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

        public ChartFormatsSequence(IStreamReader reader) : base(reader)
        {
            // Chart
            this.Chart = (Chart)OfficeGraphBiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // *2FONTLIST
            this.FontListSequences = new List<FontListSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == RecordNumber.FrtFontList)
            {
                this.FontListSequences.Add(new FontListSequence(reader));
            }
            
            // Scl
            this.Scl = (Scl)OfficeGraphBiffRecord.ReadRecord(reader);

            // PlotGrowth
            this.PlotGrowth = (PlotGrowth)OfficeGraphBiffRecord.ReadRecord(reader);

            // [FRAME]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == RecordNumber.Frame)
            {
                this.FrameSequence = new FrameSequence(reader);
            }

            // *SERIESFORMAT
            this.SeriesFormatSequences = new List<SeriesFormatSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == RecordNumber.Series)
            {
                this.SeriesFormatSequences.Add(new SeriesFormatSequence(reader));
            }

            // *SS
            this.SsSequences = new List<SsSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == RecordNumber.DataFormat)
            {
                this.SsSequences.Add(new SsSequence(reader));
            }

            // ShtProps
            this.ShtProps = (ShtProps)OfficeGraphBiffRecord.ReadRecord(reader);

            // *2DFTTEXT
            this.DftTextSequences = new List<DftTextSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == RecordNumber.DataLabExt)
            {
                this.DftTextSequences.Add(new DftTextSequence(reader));
            }

            // AxesUsed
            this.AxesUsed = (AxesUsed)OfficeGraphBiffRecord.ReadRecord(reader);


        }

    }
}
