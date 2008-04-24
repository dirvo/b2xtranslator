using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(TypeCode = 0xF004)]
    public class ShapeContainer : RegularContainer
    {
        public ShapeContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
