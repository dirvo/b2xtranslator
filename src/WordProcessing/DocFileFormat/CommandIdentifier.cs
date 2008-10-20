using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class CommandIdentifier: ByteStructure
    {
        public enum CidType
        {
            cmtFci = 0x1,
            cmtMacro = 0x2,
            cmtAllocated = 0x3,
            cmtNil = 0x7
        }

        public const int CID_LENGTH = 4;

        public Int16 ibstMacro;

        public CommandIdentifier(VirtualStreamReader reader)
            : base(reader, CID_LENGTH)
        {
            byte[] bytes = reader.ReadBytes(4);

            CidType type = (CidType)Utils.BitmaskToInt((int)bytes[0], 0x07);

            switch (type)
            {   
                case CidType.cmtFci:
                    break;
                case CidType.cmtMacro:
                    this.ibstMacro = System.BitConverter.ToInt16(bytes, 2);
                    break;
                case CidType.cmtAllocated:
                    break;
                case CidType.cmtNil:
                    break;
                default:
                    break;
            }
        }
    }
}
