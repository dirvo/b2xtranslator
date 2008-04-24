using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class GRatioAtom
    {
        public Int32 Numer;
        public Int32 Denom;

        public GRatioAtom(BinaryReader reader)
        {
            this.Numer = reader.ReadInt32();
            this.Denom = reader.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("RatioAtom({0}, {1})", this.Numer, this.Denom);
        }
    }

}
