using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(77)]
    class LeftArrowCalloutType : ShapeType
    {
        public LeftArrowCalloutType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m@0,l@0@3@2@3@2@1,,10800@2@4@2@5@0@5@0,21600,21600,21600,21600,xe"; 
            this.Formulas = new List<string>();
     
            this.Formulas.Add("val #0 ");     
            this.Formulas.Add("val #1 ");     
            this.Formulas.Add("val #2 ");     
            this.Formulas.Add("val #3 ");     
            this.Formulas.Add("sum 21600 0 #1");      
            this.Formulas.Add("sum 21600 0 #3");      
            this.Formulas.Add("sum #0 21600 0"); 

            this.AdjustmentValues = "7200,5400,3600,8100";
            this.ConnectorLocations = "@7,0;0,10800;@7,21600;21600,10800";
            this.ConnectorAngles = "270,180,90,0"; 

            this.TextboxRectangle = "@0,0,21600,21600";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="#0,topLeft";
            HandleOne.xrange="@2,21600";
            this.Handles.Add(HandleOne);

            Handle HandleTwo = new Handle();
            HandleTwo.position="topLeft,#1";
            HandleTwo.yrange="0,@3";
            this.Handles.Add(HandleTwo);

            Handle HandleThree = new Handle();
            HandleThree.position="#2,#3";
            HandleThree.xrange="0,@0";
            HandleThree.yrange = "@1,10800";
            this.Handles.Add(HandleThree); 
        }
    }
}


