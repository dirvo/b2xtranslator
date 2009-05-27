using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class PicfSequence : OfficeGraphBiffRecordSequence
    {
        public Begin Begin;

        public PicF PicF;

        public End End; 

        public PicfSequence(IStreamReader reader)
            : base(reader)
        {
            // PICF = Begin PicF End
            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // PicF 
            this.PicF = (PicF)OfficeGraphBiffRecord.ReadRecord(reader);

            // End 
            this.End = (End)OfficeGraphBiffRecord.ReadRecord(reader); 

        }
    }
}
