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
            byte[] segmentValues = readMsoArray(pSegmentInfo);
            this.Segments = new List<PathSegment>();
            for (int i = 0; i < segmentValues.Length; i += 2)
            {
                this.Segments.Add(new PathSegment(System.BitConverter.ToUInt16(segmentValues, i)));
            }

            //parse the values
            byte[] verticeValues = readMsoArray(pVertices);
            this.Values = new Int16[(verticeValues.Length / 2)];
            int j = 0;
            for (int i = 0; i < verticeValues.Length; i += 2)
            {
                this.Values[j] = System.BitConverter.ToInt16(verticeValues, i);
                j++;
            }

            // build the path
            this.VmlPath = new StringBuilder();
            int valuePointer = 0; 
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

        private byte[] readMsoArray(byte[] array)
        {
            UInt16 nElems = System.BitConverter.ToUInt16(array, 0);
            UInt16 nElemsAlloc = System.BitConverter.ToUInt16(array, 2);
            UInt16 cbElem = System.BitConverter.ToUInt16(array, 4);
            if (cbElem == 0xFFF0)
            {
                cbElem = 4;
            }
            byte[] data = new byte[cbElem * nElems];

            for (int i = 0; i < nElems; i++)
            {
                data[i] = array[6 + (i * cbElem)];
            }

            return data;
        }
    }
}
