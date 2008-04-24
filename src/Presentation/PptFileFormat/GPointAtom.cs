using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class GPointAtom
    {
        public Int32 X;
        public Int32 Y;

        public GPointAtom(BinaryReader reader)
        {
            this.X = reader.ReadInt32();
            this.Y = reader.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("PointAtom({0}, {1})", this.X, this.Y);
        }
    }

}
