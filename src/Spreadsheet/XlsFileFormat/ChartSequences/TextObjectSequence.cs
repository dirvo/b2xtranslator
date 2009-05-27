using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class TextObjectSequence : OfficeGraphBiffRecordSequence
    {
        public TxO TxO;

        public List<Continue> Continue; 

        public TextObjectSequence(IStreamReader reader)
            : base(reader)
        {
            //TEXTOBJECT = TxO *Continue
            // TxO
            this.TxO = (TxO)OfficeGraphBiffRecord.ReadRecord(reader);

            // Continue
            this.Continue = new List<Continue>(); 
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.Continue)
            {
                this.Continue.Add((Continue)OfficeGraphBiffRecord.ReadRecord(reader));
            }
        }
    }
}
