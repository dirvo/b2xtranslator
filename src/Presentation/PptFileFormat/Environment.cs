using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(1010)]
    public class Environment : RegularContainer
    {
        public Environment(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (Record rec in Children)
                {
                    switch (rec.TypeCode)
                    {
                        case 0x7d5: //FontCollectionContainer
                            break;
                        case 0xfa3: //TextMasterStyleAtom
                            TextMasterStyleAtom a = (TextMasterStyleAtom)rec;
                            break;
                        case 0xfa4: //TextCFExceptionAtom
                            TextCFExceptionAtom ce = (TextCFExceptionAtom)rec;
                            break;
                        case 0xfa5: //TextPFExceptionAtom
                            TextPFExceptionAtom e = (TextPFExceptionAtom)rec;
                            break;
                        case 0xfa9: //TextSIEExceptionAtom
                            break;
                        case 0xfc8: //KinsokuContainer
                            break;
                        default:
                            break;
                    }
                }
        
        }
    }
}
