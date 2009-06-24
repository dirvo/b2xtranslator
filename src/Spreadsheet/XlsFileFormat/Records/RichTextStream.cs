﻿using System;
using System.Text;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.RichTextStream)]
    public class RichTextStream : BiffRecord
    {
        public FrtHeader frtHeader;

        public UInt32 dwCheckSum;

        public UInt32 cb;

        public string rgb;

        public RichTextStream(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            this.frtHeader = new FrtHeader(reader);
            this.dwCheckSum = reader.ReadUInt32();
            this.cb = reader.ReadUInt32();
            Encoding codepage = Encoding.GetEncoding(1252);
            this.rgb = codepage.GetString(reader.ReadBytes((int)this.cb));
        }
    }
}
