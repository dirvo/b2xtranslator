using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(104)]
    class CurvedUpArrowType : ShapeType
    {
        public CurvedUpArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "ar0@22@3@21,,0@4@21@14@22@1@21@7@21@12@2l@13@2@8,0@11@2wa0@22@3@21@10@2@16@24@14@22@1@21@16@24@14,xewr@14@22@1@21@7@21@16@24nfe";
            this.Formulas = new List<string>();


            this.Formulas.Add("val #0");  
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("val #2"); 
            this.Formulas.Add("sum #0 width #1"); 
            this.Formulas.Add("prod @3 1 2"); 
            this.Formulas.Add("sum #1 #1 width ");
            this.Formulas.Add("sum @5 #1 #0"); 
            this.Formulas.Add("prod @6 1 2"); 
            this.Formulas.Add("mid width #0"); 
            this.Formulas.Add("ellipse #2 height @4"); 
            this.Formulas.Add("sum @4 @9 0 ");
            this.Formulas.Add("sum @10 #1 width"); 
            this.Formulas.Add("sum @7 @9 0 ");
            this.Formulas.Add("sum @11 width #0 ");
            this.Formulas.Add("sum @5 0 #0 ");
            this.Formulas.Add("prod @14 1 2 ");
            this.Formulas.Add("mid @4 @7 ");
            this.Formulas.Add("sum #0 #1 width ");
            this.Formulas.Add("prod @17 1 2 ");
            this.Formulas.Add("sum @16 0 @18 ");
            this.Formulas.Add("val width ");
            this.Formulas.Add("val height ");
            this.Formulas.Add("sum 0 0 height"); 
            this.Formulas.Add("sum @16 0 @4 ");
            this.Formulas.Add("ellipse @23 @4 height ");
            this.Formulas.Add("sum @8 128 0 ");
            this.Formulas.Add("prod @5 1 2 ");
            this.Formulas.Add("sum @5 0 128 ");
            this.Formulas.Add("sum #0 @16 @11 ");
            this.Formulas.Add("sum width 0 #0 ");
            this.Formulas.Add("prod @29 1 2 ");
            this.Formulas.Add("prod height height 1 ");
            this.Formulas.Add("prod #2 #2 1 ");
            this.Formulas.Add("sum @31 0 @32 ");
            this.Formulas.Add("sqrt @33 ");
            this.Formulas.Add("sum @34 height 0 ");
            this.Formulas.Add("prod width height @35"); 
            this.Formulas.Add("sum @36 64 0 ");
            this.Formulas.Add("prod #0 1 2 ");
            this.Formulas.Add("ellipse @30 @38 height ");
            this.Formulas.Add("sum @39 0 64 ");
            this.Formulas.Add("prod @4 1 2");
            this.Formulas.Add("sum #1 0 @41 ");
            this.Formulas.Add("prod height 4390 32768");
            this.Formulas.Add("prod height 28378 32768");

            this.AdjustmentValues = "12960,19440,7200";
            this.ConnectorLocations = "@8,0;@11,@2;@15,0;@16,@21;@13,@2";
            this.ConnectorAngles = "270,270,270,90,0";

            this.TextboxRectangle = "@41,@43,@42,@44";
           
            this.Handles = new List<Handle>();
     
            Handle HandleOne = new Handle();
            HandleOne.position="#0,topLeft";
            HandleOne.xrange="@37,@27";
            this.Handles.Add(HandleOne);

            Handle HandleTwo = new Handle();
            HandleOne.position="#1,topLeft";
            HandleOne.xrange="@25,@20";
            this.Handles.Add(HandleTwo);

            Handle HandleThree = new Handle();
            HandleThree.position="bottomRight,#2";
            HandleThree.yrange="0,@40";
            this.Handles.Add(HandleThree);
        }
    }
}
