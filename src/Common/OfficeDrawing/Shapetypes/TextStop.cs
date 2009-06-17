﻿using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(137)]
    public class TextStop : ShapeType
    {
        public TextStop()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.ExtrusionOk = true;
            this.Lock = new ProtectionBooleans();
            this.Lock.fUsefLockText = true;
            this.Lock.fLockText = true;
            this.LockShapeType = true;

            this.AdjustmentValues = "4800";
            this.Path = "m0@0l7200,r7200,l21600@0m0@1l7200,21600r7200,l21600@1e";
            this.ConnectorType = "rect";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 @0");

            this.Handles = new List<Handle>();
            Handle h1 = new Handle();
            h1.position="topLeft,#0";
            h1.yrange = "3086,10800";
            this.Handles.Add(h1);
        }
    }
}
