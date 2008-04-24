using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(TypeCode = 1000)]
    public class PptDocumentRecord : RegularContainer
    {
        public PptDocumentRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
