using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(TypeCode = 0xF002)]
    public class DrawingContainer : RegularContainer
    {
        public DrawingContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
