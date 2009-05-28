using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class BiffRecordSequence
    {
        IStreamReader _reader;
        public IStreamReader Reader
        {
            get { return _reader; }
            set { this._reader = value; }
        }

        public BiffRecordSequence(IStreamReader reader)
        {
            _reader = reader;
        }

        public BiffRecord GetNextRecord()
        {
            throw new NotImplementedException();
        }
    }
}
