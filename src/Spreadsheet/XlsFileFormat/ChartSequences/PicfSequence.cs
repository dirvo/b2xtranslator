using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords.Graph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class PicfSequence : BiffRecordSequence
    {
        public Begin Begin;

        public PicF PicF;

        public End End; 

        public PicfSequence(IStreamReader reader)
            : base(reader)
        {
            // PICF = Begin PicF End
            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // PicF 
            this.PicF = (PicF)BiffRecord.ReadRecord(reader);

            // End 
            this.End = (End)BiffRecord.ReadRecord(reader); 

        }
    }
}
