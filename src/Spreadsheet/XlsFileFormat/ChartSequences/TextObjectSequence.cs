using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class TextObjectSequence : BiffRecordSequence
    {
        public TxO TxO;

        public List<Continue> Continue; 

        public TextObjectSequence(IStreamReader reader)
            : base(reader)
        {
            //TEXTOBJECT = TxO *Continue
            // TxO
            this.TxO = (TxO)BiffRecord.ReadRecord(reader);

            // Continue
            this.Continue = new List<Continue>(); 
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continue.Add((Continue)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
