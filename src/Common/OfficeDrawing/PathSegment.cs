﻿using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class PathSegment
    {
        public enum SegmentType
        {
            msopathLineTo,
            msopathCurveTo,
            msopathMoveTo,
            msopathClose,
            msopathEnd,
            msopathEscape,
            msopathClientEscape,
            msopathInvalid
        }

        public SegmentType Type { get; set; }

        public int Count { get; set; }

        public int EscapeCode { get; set; }

        public int VertexCount { get; set; }

        public PathSegment(UInt16 segment)
        {
            this.Type = (SegmentType)Utils.BitmaskToInt(segment, 0xE000);

            if (Type == SegmentType.msopathEscape)
            {
                this.EscapeCode = Utils.BitmaskToInt(segment, 0x1F00);
                this.VertexCount = Utils.BitmaskToInt(segment, 0x00FF);
            }
            else
            {
                this.Count = Utils.BitmaskToInt(segment, 0x1FFF);
            }

        }
    }
}
