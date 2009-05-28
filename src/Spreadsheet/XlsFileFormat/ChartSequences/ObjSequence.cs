using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ObjSequence : BiffRecordSequence
    {
        public List<Continue> Continue;

        // public Obj Obj; 
 
        public ObjSequence(IStreamReader reader)
            : base(reader)
        {
            //OBJ = Obj *Continue

            // Obj
            // this.Obj = (Obj)BiffRecord.ReadRecord(reader); 

            // *Continue
            this.Continue = new List<Continue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continue.Add((Continue)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
