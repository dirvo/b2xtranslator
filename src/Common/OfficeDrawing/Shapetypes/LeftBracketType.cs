using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(85)]
    public class LeftBracketType : ShapeType
    {
        public LeftBracketType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.round;
            //Endcaps: Flat

            this.Path = "m21600,qx0@0l0@1qy21600,21600e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("prod #0 9598 32768");
            this.Formulas.Add("sum 21600 0 @2");

            this.AdjustmentValues = "1800";
            this.ConnectorLocations = "21600,0;0,10800;21600,21600";
            this.TextboxRectangle = "6326,@2,21600,@3";


            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="topLeft,#0";
            HandleOne.yrange = "0,10800";
            this.Handles.Add(HandleOne);

        }
    }
}
