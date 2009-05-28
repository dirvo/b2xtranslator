using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class WindowSequence : BiffRecordSequence
    {
        public Window2 Window2;

        public PLV PLV;

        public Scl Scl;

        public Pane Pane;

        public List<Selection> Selections;

        public WindowSequence(IStreamReader reader) 
            : base(reader)
        {
            // Window2 [PLV] [Scl] [Pane] *Selection

            // Window2
            this.Window2 = (Window2)BiffRecord.ReadRecord(reader);

            // [PLV] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PLV)
            {
                this.PLV = (PLV)BiffRecord.ReadRecord(reader);
            }
            
            // [Scl] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Scl)
            {
                this.Scl = (Scl)BiffRecord.ReadRecord(reader);
            }
            
            // [Pane] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Pane)
            {
                this.Pane = (Pane)BiffRecord.ReadRecord(reader);
            }
            
            //*Selection
            this.Selections = new List<Selection>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Selection)
            {
                this.Selections.Add((Selection)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
