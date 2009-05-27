using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ObjSequence : OfficeGraphBiffRecordSequence
    {
        public List<Continue> Continue;

        // public Obj Obj; 
 
        public ObjSequence(IStreamReader reader)
            : base(reader)
        {
            //OBJ = Obj *Continue

            // Obj
            // this.Obj = (Obj)OfficeGraphBiffRecord.ReadRecord(reader); 

            // *Continue
            this.Continue = new List<Continue>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.Continue)
            {
                this.Continue.Add((Continue)OfficeGraphBiffRecord.ReadRecord(reader));
            }
        }
    }
}
