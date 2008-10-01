using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(82)]
    class UpDownArrowCalloutType : ShapeType
    {
        public UpDownArrowCalloutType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m0@0l@3@0@3@2@1@2,10800,0@4@2@5@2@5@0,21600@0,21600@8@5@8@5@9@4@9,10800,21600@1@9@3@9@3@8,0@8xe"; 
            this.Formulas = new List<string>();

            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("val #2"); 
            this.Formulas.Add("val #3"); 
            this.Formulas.Add("sum 21600 0 #1"); 
            this.Formulas.Add("sum 21600 0 #3"); 
            this.Formulas.Add("sum #0 21600 0"); 
            this.Formulas.Add("prod @6 1 2"); 
            this.Formulas.Add("sum 21600 0 #0"); 
            this.Formulas.Add("sum 21600 0 #2");

            this.AdjustmentValues = "5400,5400,2700,8100";
            this.ConnectorLocations = "10800,0;0,10800;10800,21600;21600,10800";
            this.ConnectorAngles = "270,180,90,0"; 

            this.TextboxRectangle = "0,@0,21600,@8";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="topLeft,#0";
            HandleOne.yrange="@2,10800";
            this.Handles.Add(HandleOne);

            Handle HandleTwo = new Handle();
            HandleTwo.position="#1,topLeft";
            HandleTwo.xrange="0,@3";
            this.Handles.Add(HandleTwo);

            Handle HandleThree = new Handle();
            HandleThree.position="#3,#2";
            HandleThree.xrange="@1,10800";
            HandleThree.yrange="0,@0";
            this.Handles.Add(HandleThree);


        }
    }
}


