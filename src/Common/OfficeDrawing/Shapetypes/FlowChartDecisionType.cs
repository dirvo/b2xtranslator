using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(110)]
    public class FlowChartDecisionType :ShapeType
    {
        public FlowChartDecisionType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m10800,l,10800,10800,21600,21600,10800xe";

            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "5400,5400,16200,16200";
        }
    }
}
