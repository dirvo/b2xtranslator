using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(202)]
    public class TextboxType : ShapeType
    {
        public TextboxType()
        {
            this.Path = "m,l,21600r21600,l21600,xe";
        }
    }
}
