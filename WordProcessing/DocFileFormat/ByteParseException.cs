using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class ByteParseException : Exception
    {
        public ByteParseException() : base()
        {
        }

        public ByteParseException(string structname) : base("Cannot parse the struct " +structname+", the length of the struct doesn't match")
        {
        }
    }
}
