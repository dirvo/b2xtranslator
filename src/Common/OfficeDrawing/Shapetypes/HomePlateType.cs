using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(15)]
    class HomePlateType : ShapeType
    {
        public HomePlateType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m@0,l,,,21600@0,21600,21600,10800xe"; 
            this.Formulas = new List<string>();

            this.Formulas.Add("val #0"); 
            this.Formulas.Add("prod #0 1 2"); 

            this.AdjustmentValues = "16200";
            this.ConnectorLocations = "@1,0;0,10800;@1,21600;21600,10800";
            this.ConnectorAngles = "270,180,90,0"; 

            this.TextboxRectangle = "0,0,10800,21600;0,0,16200,21600;0,0,21600,21600";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="#0,topLeft";
            HandleOne.xrange="0,21600";


            this.Handles.Add(HandleOne);


        }
    }
}


