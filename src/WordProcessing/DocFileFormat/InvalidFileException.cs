using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class InvalidFileException :Exception
    {
        public InvalidFileException()
            : base()
        {
        }

        public InvalidFileException(string text)
            : base(text)
        {
        }
    }
}
