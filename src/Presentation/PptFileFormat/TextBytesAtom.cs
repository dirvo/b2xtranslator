using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(TypeCode = 4008)]
    public class TextBytesAtom : TextAtom
    {
        public static Encoding ENCODING = Encoding.GetEncoding("iso-8859-1");

        public TextBytesAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance, ENCODING) { }
    }

}
