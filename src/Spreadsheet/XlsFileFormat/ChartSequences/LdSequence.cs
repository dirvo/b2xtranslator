using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class LdSequence : OfficeGraphBiffRecordSequence
    {
        public Legend Legend;

        public Begin Begin;

        public Pos Pos;

        public AttachedLabelSequence AttachedLabelSequence;

        public FrameSequence FrameSequence;

        public End End; 
        

        public LdSequence(IStreamReader reader)
            : base(reader)
        {
            /// LD = Legend Begin Pos ATTACHEDLABEL [FRAME] End
            /// 

            // Legend 
            this.Legend = (Legend)OfficeGraphBiffRecord.ReadRecord(reader); 

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // Pos 
            this.Pos = (Pos)OfficeGraphBiffRecord.ReadRecord(reader);

            // [ATTACHEDLABEL]
            this.AttachedLabelSequence = new AttachedLabelSequence(reader); 

            // [FRAME]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.Frame)
            {
                this.FrameSequence = new FrameSequence(reader);
            }

            // End 
            this.End = (End)OfficeGraphBiffRecord.ReadRecord(reader); 

        }
    }
}
