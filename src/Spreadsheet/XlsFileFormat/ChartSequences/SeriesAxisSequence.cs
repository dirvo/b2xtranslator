using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class SeriesAxisSequence : OfficeGraphBiffRecordSequence
    {
        public Axis Axis; 

        public Begin Begin;

        public CatSerRange CatSerRange;

        public AxsSequence AxsSequence; 

        public End End; 

        public SeriesAxisSequence(IStreamReader reader)
            : base(reader)
        {
            // SERIESAXIS = Axis Begin [CatSerRange] AXS End 

            // Axis
            this.Axis = (Axis)OfficeGraphBiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // [CatSerRange]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.CatSerRange)
            {
                this.CatSerRange = (CatSerRange)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // AXS
            this.AxsSequence = new AxsSequence(reader); 
            
            // End 
            this.End = (End)OfficeGraphBiffRecord.ReadRecord(reader); 
        }
    }
}
