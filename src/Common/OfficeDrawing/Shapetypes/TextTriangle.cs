﻿using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(138)]
    public class TextTriangle : ShapeType
    {
        public TextTriangle()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;

            this.AdjustmentValues = "10800";
            this.Path = "m0@0l10800,,21600@0m,21600r10800,l21600,21600e";
            this.ConnectorLocations = "10800,0;5400,@1;10800,21600;16200,@1";
            this.ConnectorAngles = "270,180,90,0";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("sum @1 10800 0");
            this.Formulas.Add("sum 21600 0 @1");

            this.Handles = new List<Handle>();
            Handle h1 = new Handle();
            h1.position="topLeft,#0";
            h1.yrange = "0,21600";
            this.Handles.Add(h1);
        }
    }
}
