﻿using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Drawing;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class GD
    {
        public int sgf;
       
        public bool fCalculatedParam1;
        public bool fCalculatedParam2;
        public bool fCalculatedParam3;

        public Int16 param1;
        public Int16 param2;
        public Int16 param3;

        public GD(UInt16 flags, Int16 p1, Int16 p2, Int16 p3)
        {
            sgf = flags & 0x1FFF;
          
            fCalculatedParam1 = Tools.Utils.BitmaskToBool(flags, 0x1 << 13);
            fCalculatedParam2 = Tools.Utils.BitmaskToBool(flags, 0x1 << 14);
            fCalculatedParam3 = Tools.Utils.BitmaskToBool(flags, 0x1 << 15);

            param1 = p1;
            param2 = p2;
            param3 = p3;
        }
    }

    public class PathParser
    {
        public List<Point> Values { get; set; }

        public List<GD> Guides { get; set; }

        public List<PathSegment> Segments { get; set; }

        public UInt16 cbElemVert;

        public PathParser(byte[] pSegmentInfo, byte[] pVertices):this(pSegmentInfo,pVertices,null)
        {}

        public PathParser(byte[] pSegmentInfo, byte[] pVertices, byte[] pGuides)
        {
            this.Guides = new List<GD>();

            if (pGuides != null && pGuides.Length > 0)
            {
                UInt16 nElemsG = System.BitConverter.ToUInt16(pGuides, 0);
                UInt16 nElemsAllocG = System.BitConverter.ToUInt16(pGuides, 2);
                UInt16 cbElemG = System.BitConverter.ToUInt16(pGuides, 4);
                for (int i = 6; i < pGuides.Length; i += cbElemG)
                {
                    this.Guides.Add(new GD(System.BitConverter.ToUInt16(pGuides, i), System.BitConverter.ToInt16(pGuides, i + 2), System.BitConverter.ToInt16(pGuides, i + 4),System.BitConverter.ToInt16(pGuides, i+6)));
                }
            }


            // parse the segments
            this.Segments = new List<PathSegment>();
            if (pSegmentInfo != null && pSegmentInfo.Length > 0)
            {
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
            }

            // parse the values
            this.Values = new List<Point>();
            UInt16 nElemsVert = System.BitConverter.ToUInt16(pVertices, 0);
            UInt16 nElemsAllocVert = System.BitConverter.ToUInt16(pVertices, 2);
            cbElemVert = System.BitConverter.ToUInt16(pVertices, 4);
            if (cbElemVert == 0xfff0) cbElemVert = 4;
            int x;
            int y;
            for (int i = 6; i <= pVertices.Length - cbElemVert; i += cbElemVert)
            {
                switch(cbElemVert)
                {
                    case 4:
                        x = System.BitConverter.ToInt16(pVertices, i);

                        if (x < 0)
                        {

                        }

                        y = System.BitConverter.ToInt16(pVertices, i + cbElemVert / 2);
                        this.Values.Add(new Point(x,y));
                        break;
                    case 8:
                        x = System.BitConverter.ToInt32(pVertices, i);

                        if (x < 0)
                        {
                            if ((uint)x > 0x80000000 && (uint)x <= 0x8000007F)
                            {
                                uint index = (uint)x - 0x80000000;
                                //TODO
                            }
                        }

                        y = System.BitConverter.ToInt32(pVertices, i + cbElemVert / 2);
                        this.Values.Add(
                             new Point(x,y));
                        break;
                }
            }
        }
    }
}
