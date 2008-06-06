using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(0xF010)]
    public class ClientAnchor : Record
    {
        public byte[] bytes;

        public ClientAnchor(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.bytes = this.Reader.ReadBytes((int)this.BodySize);
        }
    }
}
