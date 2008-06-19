using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(109)]
    public class FlowChartProcessType : ShapeType
    {
        public FlowChartProcessType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m,l,21600r21600,l21600,xe";
            this.ConnectorLocations = "Rectangle";

        }
    }
}
