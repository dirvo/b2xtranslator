using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(177)]
    class FlowChartOffpageConnectorType : ShapeType
    {
        public FlowChartOffpageConnectorType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m,l21600,r,17255l10800,21600,,17255xe"; 
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "0,0,21600,17255";
        }
    }
}


