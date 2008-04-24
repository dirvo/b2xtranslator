using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class UnknownRecord : Record
    {
        public UnknownRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Reader.ReadBytes((int)size);
        }
    }

}
