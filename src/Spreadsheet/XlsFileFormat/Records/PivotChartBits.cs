using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.PivotChartBits)]
    public class PivotChartBits : BiffRecord
    {
        /// <summary>
        /// An unsigned integer that specifies the FRT record type. <br/>
        /// MUST be 0x0859.
        /// </summary>
        public UInt16 rt;

        /// <summary>
        /// A bit that specifies whether to hide the pivot field captions in the Pivot Chart.
        /// </summary>
        public bool fGXHide;

        public PivotChartBits(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            long startPos = reader.BaseStream.Position;

            this.rt = reader.ReadUInt16();
            reader.ReadBytes(2);
            UInt16 flags = reader.ReadUInt16();
            this.fGXHide = Utils.BitmaskToBool(flags, 0x1);

            reader.BaseStream.Seek(startPos, System.IO.SeekOrigin.Begin);
            reader.BaseStream.Seek(length, System.IO.SeekOrigin.Current);
        }
    }
}
