using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(TypeCode = 3999)]
    public class TextHeaderAtom : Record
    {
        public UInt32 TextType;

        public TextHeaderAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.TextType = this.Reader.ReadUInt32();
        }
    }

}
