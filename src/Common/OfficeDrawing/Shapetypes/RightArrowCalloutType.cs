using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(78)]
    class RightArrowCalloutType : ShapeType
    {
        public RightArrowCalloutType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m,l,21600@0,21600@0@5@2@5@2@4,21600,10800@2@1@2@3@0@3@0,x"; 
            this.Formulas = new List<string>();
    
            this.Formulas.Add("val #0");     
            this.Formulas.Add("val #1");     
            this.Formulas.Add("val #2");     
            this.Formulas.Add("val #3");     
            this.Formulas.Add("sum 21600 0 #1");     
            this.Formulas.Add("sum 21600 0 #3");     
            this.Formulas.Add("prod #0 1 2");

            this.AdjustmentValues = "14400,5400,18000,8100";
            this.ConnectorLocations = "@6,0;0,10800;@6,21600;21600,10800";
            this.ConnectorAngles = "270,180,90,0"; 

            this.TextboxRectangle = "0,0,@0,21600";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="#0,topLeft";
            HandleOne.xrange="0,@2";
            this.Handles.Add(HandleOne);

            Handle HandleTwo = new Handle();
            HandleTwo.position="bottomRight,#1";
            HandleTwo.yrange="0,@3";
            this.Handles.Add(HandleTwo);

            Handle HandleThree = new Handle();
            HandleThree.position="#2,#3";
            HandleThree.xrange="@0,21600";
            HandleThree.yrange = "@1,10800";
            this.Handles.Add(HandleThree);


        }
    }
}


