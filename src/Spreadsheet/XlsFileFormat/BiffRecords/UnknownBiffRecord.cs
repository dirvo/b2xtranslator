using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class UnknownBiffRecord : BiffRecord
    {
        public byte[] Content;

        public UnknownBiffRecord(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            this.Content = reader.ReadBytes((int)length);
        }
    }
}
