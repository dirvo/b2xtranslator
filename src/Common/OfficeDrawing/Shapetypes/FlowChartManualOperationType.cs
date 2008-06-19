using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(119)]
    public class FlowChartManualOperationType : ShapeType
    {
        public FlowChartManualOperationType()
        {
            this.Path = "m,l21600,,17240,21600r-12880,xe";

            this.ConnectorLocations = "10800,0;2180,10800;10800,21600;19420,10800";

            this.TextboxRectangle = "4321,0,17204,21600";

        }
    }
}
