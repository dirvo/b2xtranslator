using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class SSlideLayoutAtom
    {
        public Int32 Geom;
        public byte[] PlaceholderIds = new byte[8];

        public SSlideLayoutAtom(BinaryReader reader)
        {
            this.Geom = reader.ReadInt32();

            for (int i = 0; i < 8; i++)
                this.PlaceholderIds[i] = reader.ReadByte();
        }

        public override string ToString()
        {
            string s = String.Join(", ",
                Array.ConvertAll<byte, string>(this.PlaceholderIds,
                delegate(byte b) { return b.ToString(); }));

            return String.Format("SSlideLayoutAtom(Geom = {0}, PlaceholderIds = [{1}])",
                this.Geom, s);
        }
    }

}
