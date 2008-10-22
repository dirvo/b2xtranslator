using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(51)]
    public class AccentBorderCallout2Type : ShapeType
    {
        public AccentBorderCallout2Type()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@0@1l@2@3@4@5nfem@4,l@4,21600nfem,l21600,r,21600l,21600xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");
            this.Formulas.Add("val #2");
            this.Formulas.Add("val #3");
            this.Formulas.Add("val #4");
            this.Formulas.Add("val #5");
            this.AdjustmentValues = "-10080,24300,-3600,4050,-1800,4050";
            this.ConnectorLocations = "@0,@1;10800,0;10800,21600;0,10800;21600,10800";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="#0,#1";
            this.Handles.Add(HandleOne);

            Handle HandleTwo = new Handle();
            HandleTwo.position="#2,#3";
            this.Handles.Add(HandleTwo);

            Handle HandleThree = new Handle();
            HandleThree.position = "#4,#5";
            this.Handles.Add(HandleThree);
        }
    }
}
