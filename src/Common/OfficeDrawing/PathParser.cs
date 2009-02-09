using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Drawing;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class PathParser
    {
        public List<Point> Values { get; set; }

        public List<PathSegment> Segments { get; set; }

        public PathParser(byte[] pSegmentInfo, byte[] pVertices)
        {
            // parse the segments
            this.Segments = new List<PathSegment>();
            UInt16 nElemsSeg = System.BitConverter.ToUInt16(pSegmentInfo, 0);
            UInt16 nElemsAllocSeg = System.BitConverter.ToUInt16(pSegmentInfo, 2);
            UInt16 cbElemSeg = System.BitConverter.ToUInt16(pSegmentInfo, 4);
            for (int i = 6; i < pSegmentInfo.Length; i += 2)
            {
                this.Segments.Add(
                    new PathSegment(
                        System.BitConverter.ToUInt16(pSegmentInfo, i)
                ));
            }

            // parse the values
            this.Values = new List<Point>();
            UInt16 nElemsVert = System.BitConverter.ToUInt16(pVertices, 0);
            UInt16 nElemsAllocVert = System.BitConverter.ToUInt16(pVertices, 2);
            UInt16 cbElemVert = System.BitConverter.ToUInt16(pVertices, 4);
            for (int i = 6; i < pVertices.Length; i += 4)
            {
                this.Values.Add(
                    new Point(
                        System.BitConverter.ToInt16(pVertices, i),
                        System.BitConverter.ToInt16(pVertices, i + 2)
                ));
            }
        }

        public string BuildVmlPath()
        {
            // build the VML Path
            StringBuilder VmlPath = new StringBuilder();
            int valuePointer = 0; 
            foreach (PathSegment seg in this.Segments)
            {
                try
                {
                    switch (seg.Type)
                    {
                        case PathSegment.SegmentType.msopathCurveTo:
                            VmlPath.Append("c");
                            VmlPath.Append(this.Values[valuePointer].X);
                            VmlPath.Append(",");
                            VmlPath.Append(this.Values[valuePointer].Y);
                            VmlPath.Append(",");
                            VmlPath.Append(this.Values[valuePointer+1].X);
                            VmlPath.Append(",");
                            VmlPath.Append(this.Values[valuePointer+1].Y);
                            VmlPath.Append(",");
                            VmlPath.Append(this.Values[valuePointer + 2].X);
                            VmlPath.Append(",");
                            VmlPath.Append(this.Values[valuePointer + 2].Y);
                            valuePointer += 3;
                            break;
                        case PathSegment.SegmentType.msopathMoveTo:
                            VmlPath.Append("m");
                            VmlPath.Append(this.Values[valuePointer].X);
                            VmlPath.Append(",");
                            VmlPath.Append(this.Values[valuePointer].Y);
                            valuePointer += 1;
                            break;
                    }
                }
                catch (IndexOutOfRangeException ex)
                {
                    // Sometimes there are more Segments than available Values.
                    // Accordingly to the spec this should never happen :)
                    break;
                }
            }

            // end the path
            VmlPath.Append("e");

            return VmlPath.ToString();
        }
    }
}
