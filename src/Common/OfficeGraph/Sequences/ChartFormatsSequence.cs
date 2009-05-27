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

        public ChartFormatsSequence(IStreamReader reader) : base(reader)
        {
            // Chart
            this.Chart = (Chart)OfficeGraphBiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // *2 FONTLIST
            this.FontListSequences = new List<FontListSequence>();
            while(true)
            {
                FontListSequence fl = checkForFontList(reader);
                if(fl != null)
                {
                    this.FontListSequences.Add(fl);
                }
                else
                {
                    break;
                }
            }
            
            // Scl
            this.Scl = (Scl)OfficeGraphBiffRecord.ReadRecord(reader);

            // PlotGrowth
            this.PlotGrowth = (PlotGrowth)OfficeGraphBiffRecord.ReadRecord(reader);

            // [FRAME]
            this.FrameSequence = checkForFrameSequence(reader);
        }

        private FrameSequence checkForFrameSequence(IStreamReader reader)
        {
            FrameSequence result = null;
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == RecordNumber.Frame)
            {
                result = new FrameSequence(reader);
            }
            return result;
        }

        /// <summary>
        /// If the next record initializes a FontList sequence, the sequence is parsed and returned.<br/>
        /// The position of the stream is right after the FontList sequence.<br/>
        /// Otherwise the return value is null and the position of the stream didn't changed.
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private FontListSequence checkForFontList(IStreamReader reader)
        {
            FontListSequence result = null;
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == RecordNumber.FrtFontList)
            {
                // parse FontListSequence
                result = new FontListSequence(reader);
            }
            return result;
        }
    }
}
