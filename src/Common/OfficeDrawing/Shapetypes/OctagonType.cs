using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(10)]
    public class OctagonType : ShapeType
    {
        public OctagonType()
        {
            this.Path = "m@0,l0@0,0@2@0,21600@1,21600,21600@2,21600@0@1,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("prod @0 2929 10000");
            this.Formulas.Add("sum width 0 @3");
            this.Formulas.Add("sum height 0 @3");
            this.Formulas.Add("val width");
            this.Formulas.Add("val height");
            this.Formulas.Add("prod width 1 2");
            this.Formulas.Add("prod height 1 2");

            this.AdjustmentValues = "6326";

            this.ConnectorLocations = "@8,0;0,@9;@8,@7;@6,@9";

            this.TextboxRectangle = "0,0,21600,21600;2700,2700,18900,18900;5400,5400,16200,16200";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="#0,topLeft";
            HandleOne.switchHandle=true;
            HandleOne.xrange="0,10800";
            this.Handles.Add(HandleOne);

            this.Limo = "10800,10800";
        }
    }
}
