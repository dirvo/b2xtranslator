using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class PathParser
    {
        public Int16[] Values { get; set; }

        public List<PathSegment> Segments { get; set; }

        public StringBuilder VmlPath { get; set; }

        public PathParser(byte[] pSegmentInfo, byte[] pVertices)
        {
            //parse the segments
            this.Segments = new List<PathSegment>();
            for (int i = 0; i < pSegmentInfo.Length; i+=2)
            {
                this.Segments.Add(new PathSegment(System.BitConverter.ToUInt16(pSegmentInfo, i)));
            }

            //parse the values
            this.Values = new Int16[(pVertices.Length / 2)];
            int j = 0;
            for (int i = 0; i < pVertices.Length; i += 2)
            {
                this.Values[j] = System.BitConverter.ToInt16(pVertices, i);
                j++;
            }

            // build the path
            this.VmlPath = new StringBuilder();

            // Skip the first 3 values
            // The first 3 values are always 2 positive integers and one negative integer.
            // I don't know the real meaning of these 3 values, but the path starts always with the 4th value.
            int valuePointer = 3; 
            foreach (PathSegment seg in this.Segments)
            {
                try
                {
                    switch (seg.Type)
                    {
                        case PathSegment.SegmentType.msopathCurveTo:
                            this.VmlPath.Append("c");
                            this.VmlPath.Append(this.Values[valuePointer]);
                            this.VmlPath.Append(",");
                            valuePointer++;
                            this.VmlPath.Append(this.Values[valuePointer]);
                            this.VmlPath.Append(",");
                            valuePointer++;
                            this.VmlPath.Append(this.Values[valuePointer]);
                            this.VmlPath.Append(",");
                            valuePointer++;
                            this.VmlPath.Append(this.Values[valuePointer]);
                            this.VmlPath.Append(",");
                            valuePointer++;
                            this.VmlPath.Append(this.Values[valuePointer]);
                            this.VmlPath.Append(",");
                            valuePointer++;
                            this.VmlPath.Append(this.Values[valuePointer]);
                            valuePointer++;
                            break;
                        case PathSegment.SegmentType.msopathMoveTo:
                            this.VmlPath.Append("m");
                            this.VmlPath.Append(this.Values[valuePointer]);
                            this.VmlPath.Append(",");
                            this.VmlPath.Append(this.Values[valuePointer + 1]);
                            valuePointer += 2;
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
            this.VmlPath.Append("e");
        }
    }
}
