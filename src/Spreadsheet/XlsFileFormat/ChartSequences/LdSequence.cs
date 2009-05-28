using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords.Graph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class LdSequence : BiffRecordSequence
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
            this.Legend = (Legend)BiffRecord.ReadRecord(reader); 

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // Pos 
            this.Pos = (Pos)BiffRecord.ReadRecord(reader);

            // [ATTACHEDLABEL]
            this.AttachedLabelSequence = new AttachedLabelSequence(reader); 

            // [FRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Frame)
            {
                this.FrameSequence = new FrameSequence(reader);
            }

            // End 
            this.End = (End)BiffRecord.ReadRecord(reader); 

        }
    }
}
