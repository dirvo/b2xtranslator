using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(117)]
    public class FlowChartPreparationType: ShapeType
    {
        public FlowChartPreparationType()
        {
            this.Path = "m4353,l17214,r4386,10800l17214,21600r-12861,l,10800xe";

            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "4353,0,17214,21600";

        }
    }
}
