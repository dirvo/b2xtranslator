using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(5003)]
    public class ProgBinaryTagDataBlob : RegularContainer
    {
        public ProgBinaryTagDataBlob(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (Record rec in Children)
                {
                    switch (rec.TypeCode)
                    {
                        case 0x40d: //GridSpacing10Atom
                            break;
                        case 0x7f8: //BlipCollection9
                            break;
                        case 0x2eeb: //SlideTime10Atom
                            break;
                        case 0xfad: //TextMasterStyle9Atom
                            break;
                        default:
                            break;
                    }
                }                    
        }
    }
}
