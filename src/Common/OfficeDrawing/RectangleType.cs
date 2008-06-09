using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeShapeTypeAttribute(1)]
    public class RectangleType : ShapeType
    {
        public RectangleType()
        {
            this.Path = "m,l,21600r21600,l21600,xe";
        }
    }
}
