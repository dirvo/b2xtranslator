using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(2040)]
    public class BlipCollection9Container : RegularContainer
    {
        public BlipCollection9Container(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(2041)]
    public class BlipEntityAtom : Record
    {
        public BitmapBlip blip;

        public BlipEntityAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

            byte winBlipType = this.Reader.ReadByte();
            byte unused = this.Reader.ReadByte();

            Record rec = Record.ReadRecord(this.Reader, 0);
            blip = (BitmapBlip)rec;
        }
    }

    [OfficeRecordAttribute(4012)]
    public class StyleTextProp9Atom : Record
    {
        public StyleTextProp9Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }
}
