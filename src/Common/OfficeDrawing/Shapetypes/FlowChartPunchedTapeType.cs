using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(122)]
    class FlowChartPunchedTapeType : ShapeType
    {
        public FlowChartPunchedTapeType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m21597,19450v-225,-558,-750,-1073,-1650,-1545c18897,17605,17585,17347,16197,17260v-1500,87,-2700,345,-3787,645c11472,18377,10910,18892,10800,19450v-188,515,-750,1075,-1613,1460c8100,21210,6825,21425,5400,21597,3937,21425,2700,21210,1612,20910,675,20525,150,19965,,19450l,2147v150,558,675,1073,1612,1460c2700,3950,3937,4165,5400,4337,6825,4165,8100,3950,9187,3607v863,-387,1425,-902,1613,-1460c10910,1632,11472,1072,12410,600,13497,300,14697,85,16197,v1388,85,2700,300,3750,600c20847,1072,21372,1632,21597,2147xe"; 
            this.ConnectorLocations = "10800,2147;0,10800;10800,19450;21600,10800";
            this.TextboxRectangle = "0,4337,21600,17260";
        }
    }
}


