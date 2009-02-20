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

        public UInt16 cbElemVert;

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
            cbElemVert = System.BitConverter.ToUInt16(pVertices, 4);
            if (cbElemVert == 0xfff0) cbElemVert = 4;
            for (int i = 6; i < pVertices.Length; i += cbElemVert)
            {
                this.Values.Add(
                    new Point(
                        System.BitConverter.ToInt16(pVertices, i),
                        System.BitConverter.ToInt16(pVertices, i + cbElemVert / 2)
                ));
            }
        }
    }
}
