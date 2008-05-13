using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(0xF00B, 0xF121, 0xF122)]
    public class ShapeOptions : Record
    {
        public struct OptionEntry
        {
            public UInt16 pid;
            public bool fBid;
            public bool fComplex;
            public UInt32 op;
            public byte[] opComplex;
        }

        public OptionEntry[] Options;

        public ShapeOptions(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            long pos = this.Reader.BaseStream.Position;

            //instance is the count of properties stored in this record
            this.Options = new OptionEntry[instance];
            for (int i = 0; i < instance; i++)
            {
                OptionEntry entry = new OptionEntry();
                UInt16 flag = this.Reader.ReadUInt16();
                entry.pid = (UInt16)Utils.BitmaskToInt(flag, 0x3FFF);
                entry.fBid = Utils.BitmaskToBool(flag, 0x4000);
                entry.fComplex = Utils.BitmaskToBool(flag, 0x8000);

                if (entry.fComplex)
                {
                    Int32 length = this.Reader.ReadInt32();
                    entry.opComplex = this.Reader.ReadBytes(length);
                }
                else
                {
                    entry.op = this.Reader.ReadUInt32();
                }

                this.Options[i] = entry;
            }

            this.Reader.BaseStream.Seek(pos + size, SeekOrigin.Begin);
        }
    }
}
