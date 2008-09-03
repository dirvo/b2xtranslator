using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class SectionDescriptor : ByteStructure
    {
        public Int16 fn;
        public Int16 fnMpr;
        public Int32 fcMpr;

        /// <summary>
        /// A signed integer that specifies the position in the WordDocument Stream where a Sepx structure is located.
        /// </summary>
        public Int32 fcSepx;

        private const int SED_LENGTH = 12;

        public SectionDescriptor(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.fn = _reader.ReadInt16();
            this.fcSepx = _reader.ReadInt32();
            this.fnMpr = _reader.ReadInt16();
            this.fcMpr = _reader.ReadInt32();
        }
    }
}
