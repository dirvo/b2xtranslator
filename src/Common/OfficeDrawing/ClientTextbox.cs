using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(TypeCode = 0xF00D)]
    public class ClientTextbox : RegularContainer
    {
        public ClientTextbox(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
