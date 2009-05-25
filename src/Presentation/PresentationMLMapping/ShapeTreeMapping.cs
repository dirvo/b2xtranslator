/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Reflection;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Drawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class ShapeTreeMapping :
        AbstractOpenXmlMapping,
        IMapping<PPDrawing>
    {
        protected int _idCnt;
        protected ConversionContext _ctx;
        protected string _footertext;
        protected string _headertext;
        protected string _datetext;
        
        public PresentationMapping<RegularContainer> parentSlideMapping = null;
        public Dictionary<AnimationInfoContainer, int> animinfos = new Dictionary<AnimationInfoContainer, int>();
        public StyleTextProp9Atom ShapeStyleTextProp9Atom = null;

        private SortedDictionary<int, int> ColumnWidthsByYPos;
        private List<int> RowHeights;
        private List<ShapeContainer> Lines;

        public ShapeTreeMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public void DynamicApply(Record record)
        {
            // Call Apply(record) with dynamic dispatch (selection based on run-time type of record)
            MethodInfo method = this.GetType().GetMethod("Apply", new Type[] { record.GetType() });

            //TraceLogger.DebugInternal(method.ToString());

            try
            {
                method.Invoke(this, new Object[] { record });
            }
            catch (TargetInvocationException e)
            {
                TraceLogger.DebugInternal(e.InnerException.ToString());
                throw e.InnerException;
            }
        }

        public void Apply(PPDrawing drawing)
        {
            Apply((RegularContainer) drawing);
        }

        public void Apply(DrawingContainer drawingContainer)
        {
            GroupContainer group = drawingContainer.FirstChildWithType<GroupContainer>();
            IEnumerator<Record> iter = group.Children.GetEnumerator();
            iter.MoveNext();

            ShapeContainer header = iter.Current as ShapeContainer;
            WriteGroupShapeProperties(header);

            while (iter.MoveNext())
                DynamicApply(iter.Current);
        }

        //used to give each group a unique identifyer
        private int groupcounter = -10;
        public void Apply(GroupContainer group)
        {
            GroupShapeRecord gsr = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>();
            ClientAnchor anchor = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();

            foreach (ShapeOptions ops in group.FirstChildWithType<ShapeContainer>().AllChildrenWithType<ShapeOptions>())
            {
                if (ops.OptionsByID.ContainsKey(ShapeOptions.PropertyId.tableProperties))
                {
                    uint TABLEFLAGS = ops.OptionsByID[ShapeOptions.PropertyId.tableProperties].op;
                    if (Tools.Utils.BitmaskToBool(TABLEFLAGS, 0x1))
                    {
                        //this group is a table
                        ApplyTable(group, TABLEFLAGS);
                        return;
                    }
                }
            }

            _writer.WriteStartElement("p", "grpSp", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "nvGrpSpPr", OpenXmlNamespaces.PresentationML);
            WriteCNvPr(--groupcounter, "");
            _writer.WriteElementString("p", "cNvGrpSpPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteEndElement(); //nvGrpSpPr

            _writer.WriteStartElement("p", "grpSpPr", OpenXmlNamespaces.PresentationML);

            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

                _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "chOff", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(gsr.rcgBounds.Left).ToString());
                _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(gsr.rcgBounds.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "chExt", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(gsr.rcgBounds.Right - gsr.rcgBounds.Left).ToString());
                _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(gsr.rcgBounds.Bottom - gsr.rcgBounds.Top).ToString());
                _writer.WriteEndElement();       

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); //grpSpPr

            foreach (Record record in group.Children)
            {
                DynamicApply(record);
            }

            _writer.WriteEndElement(); //grpSp
        }

        private int GetGridSpanCount(ChildAnchor anch, int col)
        {
            int count = 0;
            int availableWidth = anch.rcgBounds.Width;
            int currentLeft = anch.Left;

            while (availableWidth > 0)
            {
                availableWidth -= ColumnWidthsByYPos[currentLeft];
                currentLeft += ColumnWidthsByYPos[currentLeft];
                count++;
            }            

            return count;
        }

        private int GetRowSpanCount(ChildAnchor anch, int row)
        {
            int count = 0;
            int availableHeight = anch.rcgBounds.Height;

            while (availableHeight > 0)
            {
                availableHeight -= RowHeights[row + count];
                count++;
            }

            return count;
        }

        public void ApplyTable(GroupContainer group, uint TABLEFLAGS)
        {
            GroupShapeRecord gsr = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>();
            ClientAnchor anchor = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();

            RowHeights = new List<int>();
            foreach (ShapeOptions ops in group.FirstChildWithType<ShapeContainer>().AllChildrenWithType<ShapeOptions>())
            {
                if (ops.OptionsByID.ContainsKey(ShapeOptions.PropertyId.tableRowProperties))
                {
                    uint TableRowPropertiesCount = ops.OptionsByID[ShapeOptions.PropertyId.tableRowProperties].op;
                    byte[] data = ops.OptionsByID[ShapeOptions.PropertyId.tableRowProperties].opComplex;
                    UInt16 nElems = BitConverter.ToUInt16(data, 0);
                    UInt16 nElemsAlloc = BitConverter.ToUInt16(data, 2);
                    UInt16 cbElem = BitConverter.ToUInt16(data, 4);
                    for (int i = 0; i < nElems; i++)
                    {
                        Int32 height = BitConverter.ToInt32(data, 6 + i * cbElem);
                        RowHeights.Add(height);
                    }
                }
            }

            List<ShapeContainer> Cells = new List<ShapeContainer>();
            Lines = new List<ShapeContainer>();
            Dictionary<string, ShapeContainer> LinesByPosition = new Dictionary<string, ShapeContainer>();
           
            SortedList<int, SortedList<int, ShapeContainer>> tablelist = new SortedList<int,SortedList<int,ShapeContainer>>();
            foreach (ShapeContainer scontainer in group.AllChildrenWithType<ShapeContainer>())
            {
                ChildAnchor anch = scontainer.FirstChildWithType<ChildAnchor>();
                foreach (Shape shape in scontainer.AllChildrenWithType<Shape>())
                {
                    if (Utils.getPrstForShape(shape.Instance) == "rect")
                    {
                        if (!tablelist.ContainsKey(anch.Top)) tablelist.Add(anch.Top, new SortedList<int, ShapeContainer>());
                        tablelist[anch.Top].Add(anch.Left, scontainer);
                    }
                    else if (Utils.getPrstForShape(shape.Instance) == "line")
                    {
                        Lines.Add(scontainer);
                    }
                } 
            }

            ColumnWidthsByYPos = new SortedDictionary<int, int>();
            Dictionary<int, int> ColumnIndices = new Dictionary<int, int>(); //this list will contain all column limits
            foreach (SortedList<int, ShapeContainer> rowlist in tablelist.Values)
            {
                //rowlist contains all cells in a row
                foreach (int y in rowlist.Keys)
                {
                    int w = rowlist[y].FirstChildWithType<ChildAnchor>().rcgBounds.Width;
                    if (!ColumnWidthsByYPos.ContainsKey(y))
                    {
                        ColumnWidthsByYPos.Add(y, w);
                    }
                    else
                    {
                        if (w < ColumnWidthsByYPos[y]) ColumnWidthsByYPos[y] = w;
                    }
                    Cells.Add(rowlist[y]);
                }
            }
            int counter = 0;
            foreach (int y in ColumnWidthsByYPos.Keys)
            {
                ColumnIndices.Add(y, counter++);
            }            

            //the table contains all cells at their correct position
            ShapeContainer[,] table = new ShapeContainer[RowHeights.Count, ColumnWidthsByYPos.Count];
            int c;
            int r = 0;
            foreach (SortedList<int, ShapeContainer> rowlst in tablelist.Values)
            {
                foreach (int y in rowlst.Keys)
                {
                    c = ColumnIndices[y];
                    table[r, c] = rowlst[y];
                }
                r++;
            }
           
            _writer.WriteStartElement("p", "graphicFrame", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "nvGraphicFramePr", OpenXmlNamespaces.PresentationML);
            WriteCNvPr(--groupcounter, "");
            _writer.WriteStartElement("p", "cNvGraphicFramePr", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("a", "graphicFrameLocks", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("noGrp", "1");
            _writer.WriteEndElement(); //graphicFrameLocks
            _writer.WriteEndElement(); //cNvGraphicFramePr
            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteEndElement(); //nvGraphicFramePr            

            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                _writer.WriteStartElement("p", "xfrm", OpenXmlNamespaces.PresentationML);

                _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteEndElement();
            }

            _writer.WriteStartElement("a", "graphic", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "graphicData", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");
            _writer.WriteStartElement("a", "tbl", OpenXmlNamespaces.DrawingML);
            _writer.WriteElementString("a", "tblPr", OpenXmlNamespaces.DrawingML, "");
            _writer.WriteStartElement("a", "tblGrid", OpenXmlNamespaces.DrawingML);

            foreach (int y in ColumnWidthsByYPos.Keys)
            {
                 _writer.WriteStartElement("a", "gridCol", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("w", Utils.MasterCoordToEMU(ColumnWidthsByYPos[y]).ToString());
                _writer.WriteEndElement(); //gridCol
            }
            _writer.WriteEndElement(); //tblGrid

            for (int row = 0; row < RowHeights.Count; row++)
            {
                _writer.WriteStartElement("a", "tr", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("h", Utils.MasterCoordToEMU(RowHeights[row]).ToString());

                for (int col = 0; col < ColumnWidthsByYPos.Count; col++)
                {
                    _writer.WriteStartElement("a", "tc", OpenXmlNamespaces.DrawingML);

                    if (table[row, col] != null)
                    {
                        ShapeContainer container = table[row, col];

                        ChildAnchor anch = container.FirstChildWithType<ChildAnchor>();
                        int colWidth = ColumnWidthsByYPos[anch.Left];
                        
                        if (anch.rcgBounds.Height > RowHeights[row])
                        {
                            _writer.WriteAttributeString("rowSpan", GetRowSpanCount(anch, row).ToString());
                        }
                        
                        if (anch.rcgBounds.Width > colWidth)
                        {
                            _writer.WriteAttributeString("gridSpan", GetGridSpanCount(anch, col).ToString());
                        }
                                              
                        foreach (Record record in container.Children)
                        {
                            if (record is ClientTextbox)
                            {
                                so = container.FirstChildWithType<ShapeOptions>();
                                Apply((ClientTextbox)record, true);
                            }
                        }

                        _writer.WriteStartElement("a", "tcPr", OpenXmlNamespaces.DrawingML);

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.anchorText))
                        {
                            switch (so.OptionsByID[ShapeOptions.PropertyId.anchorText].op)
                            {
                                case 0: //Top
                                    _writer.WriteAttributeString("anchor", "t");
                                    break;
                                case 1: //Middle
                                    _writer.WriteAttributeString("anchor", "ctr");
                                    break;
                                case 2: //Bottom
                                    _writer.WriteAttributeString("anchor", "b");
                                    break;
                                case 3: //TopCentered
                                    _writer.WriteAttributeString("anchor", "t");
                                    _writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 4: //MiddleCentered
                                    _writer.WriteAttributeString("anchor", "ctr");
                                    _writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 5: //BottomCentered
                                    _writer.WriteAttributeString("anchor", "b");
                                    _writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                            }
                        }

                        _writer.WriteAttributeString("horzOverflow", "overflow");

                        WriteTableLineProperties(row, col, table);

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
                        {
                            FillStyleBooleanProperties p = new FillStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                            if (p.fUsefFilled & p.fFilled) //  so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                            {
                                new FillMapping(_ctx, _writer, parentSlideMapping).Apply(so);
                            }
                        }


                        _writer.WriteEndElement(); //tcPr

                    }
                    else
                    {
                        if (col > 0 && table[row, col - 1] != null)
                        {
                            ShapeContainer previouscontainer = table[row, col-1];
                            ChildAnchor anch = previouscontainer.FirstChildWithType<ChildAnchor>();
                            if (anch.rcgBounds.Height > RowHeights[row])
                            {
                                _writer.WriteAttributeString("rowSpan", "2");                                
                            }

                            if (anch.rcgBounds.Width > ColumnWidthsByYPos[anch.Left])
                            {
                                _writer.WriteAttributeString("hMerge", "1");
                            }
                        }
                        else if (row > 0 && col > 0 && table[row - 1, col - 1] != null) //this checks the cell above on the left
                        {
                            ShapeContainer previouscontainer = table[row - 1, col - 1];
                            ChildAnchor anch = previouscontainer.FirstChildWithType<ChildAnchor>();

                            if (anch.rcgBounds.Width > ColumnWidthsByYPos[anch.Left])
                            {
                                _writer.WriteAttributeString("hMerge", "1");
                            }
                        }

                        if (row > 0 && table[row-1, col] != null) //this checks the cell on the left
                        {
                            ShapeContainer previouscontainer = table[row-1, col];
                            ChildAnchor anch = previouscontainer.FirstChildWithType<ChildAnchor>();
                            int colWidth = ColumnWidthsByYPos[anch.Left];

                            if (anch.rcgBounds.Width > colWidth)
                            {
                                _writer.WriteAttributeString("gridSpan", "2");
                            }

                            if (anch.rcgBounds.Height > RowHeights[row-1])
                            {
                                _writer.WriteAttributeString("vMerge", "1");
                            }
                           
                        }
                        else if (row > 0 && col > 0  && table[row - 1, col - 1] != null) //this checks the cell above on the left
                        {
                            ShapeContainer previouscontainer = table[row - 1, col-1];
                            ChildAnchor anch = previouscontainer.FirstChildWithType<ChildAnchor>();

                            if (anch.rcgBounds.Height > RowHeights[row - 1])
                            {
                                _writer.WriteAttributeString("vMerge", "1");
                            }
                        }

                        //insert dummy tc content
                        _writer.WriteStartElement("a","txBody",OpenXmlNamespaces.DrawingML);
                        _writer.WriteElementString("a","bodyPr",OpenXmlNamespaces.DrawingML,"");
                        _writer.WriteElementString("a","lstStyle",OpenXmlNamespaces.DrawingML,"");
                        _writer.WriteElementString("a","p",OpenXmlNamespaces.DrawingML,"");
                        _writer.WriteEndElement(); //txBody

                        _writer.WriteElementString("a", "tcPr", OpenXmlNamespaces.DrawingML,"");
                    }

                    _writer.WriteEndElement(); //tc

                    //cellPointer++;
                }


                _writer.WriteEndElement(); //tr
            }

            _writer.WriteEndElement(); //tbl
            _writer.WriteEndElement(); //graphicData
            _writer.WriteEndElement(); //graphic
            _writer.WriteEndElement(); //graphicFrame

        }

        private void WriteTableLineProperties(int row, int col, ShapeContainer[,] table)
        {
            int rows = table.GetLength(0);
            int columns = table.GetLength(1);

            //check cell position
            bool isLeft = (col == 0);
            bool isRight = (col == columns - 1);
            bool isTop = (row == 0);
            bool isBottom = (row == rows - 1);

            int span;

            ShapeContainer container = table[row, col];
            ChildAnchor anch = container.FirstChildWithType<ChildAnchor>();
            int colWidth = ColumnWidthsByYPos[anch.Left];
            if (anch.rcgBounds.Height > RowHeights[row])
            {
                //recheck isBottom
                span = GetRowSpanCount(anch, row);
                isBottom = (row + span - 1 == rows - 1);
            }
            if (anch.rcgBounds.Width > colWidth)
            {
                //recheck isRight
                span = GetGridSpanCount(anch, col);
                isRight = (col + span - 1 == columns - 1);
            }
            

            ShapeContainer outerLine = null;
            ShapeContainer innerLine = null;

            int topmost = -1;
            int topleast = int.MaxValue;
            foreach (ShapeContainer lineCont in Lines)
            {
                //first find topmost and topleast vertical lines
                //the topmost line is used as outer line
                foreach (ChildAnchor anchor in lineCont.AllChildrenWithType<ChildAnchor>())
                {
                    //vertical
                    if (anchor.Top == anchor.Bottom)
                    {
                        if (anchor.Top > topmost)
                        {
                            outerLine = lineCont;
                            topmost = anchor.Top;
                        }
                        if (anchor.Top < topleast)
                        {
                            topleast = anchor.Top;
                        }
                    }
                }
            }
            foreach (ShapeContainer lineCont in Lines)
            {
                //a vertical line which is neither topmost nor topleast is used as inner line
                foreach (ChildAnchor anchor in lineCont.AllChildrenWithType<ChildAnchor>())
                {
                    //vertical
                    if (anchor.Top == anchor.Bottom)
                    {
                        if (anchor.Top < topmost & anchor.Top > topleast)
                        {
                            innerLine = lineCont;
                            break;
                        }
                    }
                }   
            }
            
     
            _writer.WriteStartElement("a", "lnL", OpenXmlNamespaces.DrawingML);
            if (isLeft) WriteLineProperties(outerLine); else WriteLineProperties(innerLine); 
            _writer.WriteEndElement(); //lnL

            _writer.WriteStartElement("a", "lnR", OpenXmlNamespaces.DrawingML);
            if (isRight) WriteLineProperties(outerLine); else WriteLineProperties(innerLine);
            _writer.WriteEndElement(); //lnR

            _writer.WriteStartElement("a", "lnT", OpenXmlNamespaces.DrawingML);
            if (isTop) WriteLineProperties(outerLine); else WriteLineProperties(innerLine);
            _writer.WriteEndElement(); //lnT

            _writer.WriteStartElement("a", "lnB", OpenXmlNamespaces.DrawingML);
            if (isBottom) WriteLineProperties(outerLine); else WriteLineProperties(innerLine);
            _writer.WriteEndElement(); //lnB

            _writer.WriteStartElement("a", "lnTlToBr", OpenXmlNamespaces.DrawingML);
            _writer.WriteElementString("a", "noFill", OpenXmlNamespaces.DrawingML, "");
            _writer.WriteEndElement(); //lnTlToBr

            _writer.WriteStartElement("a", "lnBlToTr", OpenXmlNamespaces.DrawingML);
            _writer.WriteElementString("a", "noFill", OpenXmlNamespaces.DrawingML, "");
            _writer.WriteEndElement(); //lnBlToTr

        }

        private void WriteLineProperties(ShapeContainer lineCont)
        {
            foreach (ShapeOptions soline in lineCont.AllChildrenWithType<ShapeOptions>())
            {
                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineWidth))
                {
                    uint w = soline.OptionsByID[ShapeOptions.PropertyId.lineWidth].op;
                    _writer.WriteAttributeString("w", w.ToString());
                }
                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndCapStyle))
                {
                    //switch (soline.OptionsByID[ShapeOptions.PropertyId.lineEndCapStyle].op)
                    //{
                    //    case 0: //round
                    //        _writer.WriteAttributeString("cap", "rnd");
                    //        break;
                    //    case 1: //square
                    //        _writer.WriteAttributeString("cap", "sq");
                    //        break;
                    //    case 2: //flat
                            _writer.WriteAttributeString("cap", "flat");
                            //break;
                    //}
                }

                _writer.WriteAttributeString("cmpd", "sng");
                _writer.WriteAttributeString("algn", "ctr");

                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                {
                    _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(soline.OptionsByID[ShapeOptions.PropertyId.lineColor].op, lineCont.FirstAncestorWithType<Slide>(), so));
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                }

                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineDashing))
                {
                    _writer.WriteStartElement("a", "prstDash", OpenXmlNamespaces.DrawingML);
                    switch ((ShapeOptions.LineDashing)soline.OptionsByID[ShapeOptions.PropertyId.lineDashing].op)
                    {
                        case ShapeOptions.LineDashing.Solid:
                            _writer.WriteAttributeString("val", "solid");
                            break;
                        case ShapeOptions.LineDashing.DashSys:
                            _writer.WriteAttributeString("val", "sysDash");
                            break;
                        case ShapeOptions.LineDashing.DotSys:
                            _writer.WriteAttributeString("val", "sysDot");
                            break;
                        case ShapeOptions.LineDashing.DashDotSys:
                            _writer.WriteAttributeString("val", "sysDashDot");
                            break;
                        case ShapeOptions.LineDashing.DashDotDotSys:
                            _writer.WriteAttributeString("val", "sysDashDotDot");
                            break;
                        case ShapeOptions.LineDashing.DotGEL:
                            _writer.WriteAttributeString("val", "dot");
                            break;
                        case ShapeOptions.LineDashing.DashGEL:
                            _writer.WriteAttributeString("val", "dash");
                            break;
                        case ShapeOptions.LineDashing.LongDashGEL:
                            _writer.WriteAttributeString("val", "lgDash");
                            break;
                        case ShapeOptions.LineDashing.DashDotGEL:
                            _writer.WriteAttributeString("val", "dashDot");
                            break;
                        case ShapeOptions.LineDashing.LongDashDotGEL:
                            _writer.WriteAttributeString("val", "lgDashDot");
                            break;
                        case ShapeOptions.LineDashing.LongDashDotDotGEL:
                            _writer.WriteAttributeString("val", "lgDashDotDot");
                            break;
                    }
                    _writer.WriteEndElement();
                }

                _writer.WriteElementString("a", "round", OpenXmlNamespaces.DrawingML, "");

                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStartArrowhead))
                {
                    ShapeOptions.LineEnd val = (ShapeOptions.LineEnd)soline.OptionsByID[ShapeOptions.PropertyId.lineStartArrowhead].op;
                    if (val != ShapeOptions.LineEnd.NoEnd)
                    {
                        _writer.WriteStartElement("a", "headEnd", OpenXmlNamespaces.DrawingML);
                        switch (val)
                        {
                            case ShapeOptions.LineEnd.ArrowEnd:
                                _writer.WriteAttributeString("type", "triangle");
                                break;
                            case ShapeOptions.LineEnd.ArrowStealthEnd:
                                _writer.WriteAttributeString("type", "stealth");
                                break;
                            case ShapeOptions.LineEnd.ArrowDiamondEnd:
                                _writer.WriteAttributeString("type", "diamond");
                                break;
                            case ShapeOptions.LineEnd.ArrowOvalEnd:
                                _writer.WriteAttributeString("type", "oval");
                                break;
                            case ShapeOptions.LineEnd.ArrowOpenEnd:
                                _writer.WriteAttributeString("type", "arrow");
                                break;
                            case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                            case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                                _writer.WriteAttributeString("type", "triangle");
                                break;
                        }
                        _writer.WriteAttributeString("w", "med");
                        _writer.WriteAttributeString("len", "med");
                        _writer.WriteEndElement(); //headEnd
                    }
                }
                else
                {
                    _writer.WriteStartElement("a", "headEnd", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("type", "none");
                    _writer.WriteAttributeString("w", "med");
                    _writer.WriteAttributeString("len", "med");
                    _writer.WriteEndElement(); //headEnd
                }

                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndArrowhead))
                {
                    ShapeOptions.LineEnd val = (ShapeOptions.LineEnd)soline.OptionsByID[ShapeOptions.PropertyId.lineEndArrowhead].op;
                    if (val != ShapeOptions.LineEnd.NoEnd)
                    {
                        _writer.WriteStartElement("a", "tailEnd", OpenXmlNamespaces.DrawingML);
                        switch (val)
                        {
                            case ShapeOptions.LineEnd.ArrowEnd:
                                _writer.WriteAttributeString("type", "triangle");
                                break;
                            case ShapeOptions.LineEnd.ArrowStealthEnd:
                                _writer.WriteAttributeString("type", "stealth");
                                break;
                            case ShapeOptions.LineEnd.ArrowDiamondEnd:
                                _writer.WriteAttributeString("type", "diamond");
                                break;
                            case ShapeOptions.LineEnd.ArrowOvalEnd:
                                _writer.WriteAttributeString("type", "oval");
                                break;
                            case ShapeOptions.LineEnd.ArrowOpenEnd:
                                _writer.WriteAttributeString("type", "arrow");
                                break;
                            case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                            case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                                _writer.WriteAttributeString("type", "triangle");
                                break;
                        }
                        _writer.WriteAttributeString("w", "med");
                        _writer.WriteAttributeString("len", "med");
                        _writer.WriteEndElement(); //tailnd
                    }
                }
                else
                {
                    _writer.WriteStartElement("a", "tailEnd", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("type", "none");
                    _writer.WriteAttributeString("w", "med");
                    _writer.WriteAttributeString("len", "med");
                    _writer.WriteEndElement(); //tailEnd
                }

                break;  
            }
        }


        private ShapeOptions so;
        public void Apply(ShapeContainer container)
        {
            Apply(container, "","", "");
        }
        public void Apply(ShapeContainer container, string footertext, string headertext, string datetext)
        {
            
            _footertext = footertext;
            _headertext = headertext;
            _datetext = datetext;
            ClientData clientData = container.FirstChildWithType<ClientData>();

            RegularContainer slide = container.FirstAncestorWithType<Slide>();
            if (slide == null) slide = container.FirstAncestorWithType<Note>();
            if (slide == null) slide = container.FirstAncestorWithType<Handout>();

            bool continueShape = true;

            Shape sh = container.FirstChildWithType<Shape>();

            so = container.FirstChildWithType<ShapeOptions>();
            //if (clientData == null)
                if (so != null)
                {
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.Pib))
                    {
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.BlipBooleanProperties))
                        {
                            BlipBooleanProperties props = new BlipBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.BlipBooleanProperties].op);
                            if (props.fPictureActive && props.fusefPictureActive)
                            {
                                writeOle(container);
                                continueShape = false;
                            }
                        }

                        if (continueShape)
                        {
                            writePic(container);
                            continueShape = false;
                        }
                    }
                }
                else
                {
                    so =  new ShapeOptions();
                }

            ShapeOptions sndSo = null;
            if (container.AllChildrenWithType<ShapeOptions>().Count > 1)
            {
                sndSo = ((RegularContainer)sh.ParentRecord).AllChildrenWithType<ShapeOptions>()[1];
                //if (sndSo.OptionsByID.ContainsKey(ShapeOptions.PropertyId.metroBlob))
                //{
                //    ZipUtils.ZipReader reader = null;
                //    try
                //    {
                //        ShapeOptions.OptionEntry metroBlob = sndSo.OptionsByID[ShapeOptions.PropertyId.metroBlob];
                //        byte[] code = metroBlob.opComplex;
                //        string path = System.IO.Path.GetTempFileName();
                //        System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Create);
                //        fs.Write(code, 0, code.Length);
                //        fs.Close();

                //        reader = ZipUtils.ZipFactory.OpenArchive(path);
                //        System.IO.StreamReader mems = new System.IO.StreamReader(reader.GetEntry("drs/shapexml.xml"));
                //        string xml = mems.ReadToEnd();
                //        xml = Tools.Utils.replaceOutdatedNamespaces(xml);
                //        xml = xml.Substring(xml.IndexOf("<p:sp")); //remove xml declaration

                //        _writer.WriteRaw(xml);

                //        continueShape = false;

                //        reader.Close();
                //    }
                //    catch (Exception e)
                //    {
                //        continueShape = true;
                //        if (reader != null) reader.Close();
                //    }
                //}
            }

            if (continueShape)
            {
                if (sh.fConnector)
                {
                    string idStart = "";
                    string idEnd = "";
                    string idxStart = "0";
                    string idxEnd = "0";
                    foreach (FConnectorRule rule in container.FirstAncestorWithType<DrawingContainer>().FirstChildWithType<SolverContainer>().AllChildrenWithType<FConnectorRule>())
                    {
                        if (rule.spidC == sh.spid) //spidC marks the connector shape
                        {
                            if (!spidToId.ContainsKey((int)rule.spidA))
                            {
                                spidToId.Add((int)rule.spidA, ++_idCnt);
                            }
                            if (!spidToId.ContainsKey((int)rule.spidB))
                            {
                                spidToId.Add((int)rule.spidB, ++_idCnt);
                            }

                            idStart = spidToId[(int)rule.spidA].ToString(); //spidA marks the start shape
                            idEnd = spidToId[(int)rule.spidB].ToString(); //spidB marks the end shape
                            idxStart = rule.cptiA.ToString();
                            idxEnd = rule.cptiB.ToString();
                        }
                    }

                    _writer.WriteStartElement("p", "cxnSp", OpenXmlNamespaces.PresentationML);
                    _writer.WriteStartElement("p", "nvCxnSpPr", OpenXmlNamespaces.PresentationML);
                    WriteCNvPr(sh.spid, "");

                    _writer.WriteStartElement("p", "cNvCxnSpPr", OpenXmlNamespaces.PresentationML);

                    _writer.WriteStartElement("a", "stCxn", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("id", idStart);
                    _writer.WriteAttributeString("idx", idxStart);
                    _writer.WriteEndElement(); //stCxn

                    _writer.WriteStartElement("a", "endCxn", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("id", idEnd);
                    _writer.WriteAttributeString("idx", idxEnd);
                    _writer.WriteEndElement(); //endCxn

                    _writer.WriteEndElement(); //cNvCxnSpPr
                }
                else
                {
                    _writer.WriteStartElement("p", "sp", OpenXmlNamespaces.PresentationML);
                    _writer.WriteStartElement("p", "nvSpPr", OpenXmlNamespaces.PresentationML);
                    WriteCNvPr(sh.spid, "");

                    _writer.WriteElementString("p", "cNvSpPr", OpenXmlNamespaces.PresentationML, "");
                }               
               
                _writer.WriteStartElement("p", "nvPr", OpenXmlNamespaces.PresentationML);

                OEPlaceHolderAtom placeholder = null;
                int exObjIdRef = 0;
                CheckClientData(clientData, ref placeholder, ref exObjIdRef);
                
                _writer.WriteEndElement();

                _writer.WriteEndElement();


                // Visible shape properties
                _writer.WriteStartElement("p", "spPr", OpenXmlNamespaces.PresentationML);

                ClientAnchor anchor = container.FirstChildWithType<ClientAnchor>();
                ChildAnchor chAnchor = container.FirstChildWithType<ChildAnchor>();

                bool swapHeightWidth = false;
                Double dc = 0;
                if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
                {
                    _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);
                    if (sh.fFlipH) _writer.WriteAttributeString("flipH", "1");
                    if (sh.fFlipV) _writer.WriteAttributeString("flipV", "1");
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.rotation))
                    {
                        byte[] bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.rotation].op);
                        int integral = BitConverter.ToInt16(bytes, 2);
                        uint fractional = BitConverter.ToUInt16(bytes, 0);
                        Decimal result = integral +((decimal)fractional / (decimal)65536);

                        Double w = anchor.Bottom - anchor.Top;
                        Double h = anchor.Right - anchor.Left;

                        dc = (w - h) / 2;

                        if (Math.Abs(result) > 45 && Math.Abs(result) < 135) swapHeightWidth = true;
                        if (Math.Abs(result) > 225 && Math.Abs(result) < 315) swapHeightWidth = true;

                        if (result < 0 && sh.fFlipH == false) result = result * -1;

                        string rotation = Math.Floor(result * 60000).ToString();
                        if (result != 0)
                        {
                            _writer.WriteAttributeString("rot", rotation);
                        }

                    }

                    if (container.FirstAncestorWithType<GroupContainer>().FirstAncestorWithType<GroupContainer>() == null)
                    {
                        _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);

                        if (swapHeightWidth)
                        {
                            _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left - (int)dc).ToString());
                            _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top + (int)dc).ToString());
                        }
                        else
                        {
                            _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                            _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                        }
                        _writer.WriteEndElement();

                        _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);

                        if (swapHeightWidth)
                        {
                            _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                            _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());                            
                        }
                        else
                        {
                            _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                            _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                        }
                        _writer.WriteEndElement();
                    }
                    else
                    {
                        _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                        _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                        _writer.WriteEndElement();

                        _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                        if (swapHeightWidth)
                        {
                            _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                            _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());                            
                        }
                        else
                        {
                            _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                            _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                        }         
                        _writer.WriteEndElement();
                    }
                    _writer.WriteEndElement();
                }
                else if (chAnchor != null && chAnchor.Right >= chAnchor.Left && chAnchor.Bottom >= chAnchor.Top)
                {
                    ClientAnchor groupAnchor = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();
                    Rectangle rec = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>().rcgBounds;

                    _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.rotation))
                    {
                        byte[] bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.rotation].op);
                        int integral = BitConverter.ToInt16(bytes, 2);
                        uint fractional = BitConverter.ToUInt16(bytes, 0);
                        Decimal result = integral + ((decimal)fractional / (decimal)65536);

                        if (result < 0 && sh.fFlipH == false) result = result * -1;

                        Double w = chAnchor.Bottom - chAnchor.Top;
                        Double h = chAnchor.Right - chAnchor.Left;

                        dc = (w - h) / 2;

                        if (Math.Abs(result) > 45 && Math.Abs(result) < 135) swapHeightWidth = true;
                        if (Math.Abs(result) > 225 && Math.Abs(result) < 315) swapHeightWidth = true;

                        string rotation = Math.Floor(result * 60000).ToString();
                        _writer.WriteAttributeString("rot", rotation);
                    }

                    if (sh.fFlipH) _writer.WriteAttributeString("flipH", "1");
                    if (sh.fFlipV) _writer.WriteAttributeString("flipV", "1");

                    _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(chAnchor.Left).ToString());
                    _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(chAnchor.Top).ToString());
                    _writer.WriteEndElement();

                    _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(chAnchor.Right - chAnchor.Left).ToString());
                    _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(chAnchor.Bottom - chAnchor.Top).ToString());
                    _writer.WriteEndElement();

                    _writer.WriteEndElement();
                }

                if (sh.Instance != 0 & !so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pSegmentInfo)) //this means a predefined shape
                {
                    WriteprstGeom(sh);
                }
                else //this means a custom shape
                {
                    WritecustGeom(sh);
                }

                WriteShapeProperties(sh, placeholder != null, slide);

                _writer.WriteEndElement();

                bool TextBoxFound = false;

                // Descend into unsupported records
                foreach (Record record in container.Children)
                {
                    DynamicApply(record);
                    if (record is ClientTextbox) TextBoxFound = true;
                }

                
                if (!TextBoxFound & !sh.fConnector)
                {

                    _writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);
                    writeBodyPr(container);
                    _writer.WriteElementString("a", "lstStyle", OpenXmlNamespaces.DrawingML, "");
                    _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

                    //check if there is text in so
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.gtextUNICODE))
                    {
                        byte[] bytes = so.OptionsByID[ShapeOptions.PropertyId.gtextUNICODE].opComplex;
                        string sText = Encoding.Unicode.GetString(bytes);
                        if (sText.Contains("\0")) sText = sText.Substring(0, sText.IndexOf("\0"));
                        _writer.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingML);

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.gtextFont))
                        {
                            bytes = so.OptionsByID[ShapeOptions.PropertyId.gtextFont].opComplex;
                            string sFont = Encoding.Unicode.GetString(bytes);
                            if (sFont.Contains("\0")) sFont = sFont.Substring(0, sFont.IndexOf("\0"));

                            _writer.WriteStartElement("a", "rPr", OpenXmlNamespaces.DrawingML);
                            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.gtextSize))
                            {
                                _writer.WriteAttributeString("sz", (so.OptionsByID[ShapeOptions.PropertyId.gtextSize].op / 0x100).ToString());
                            }
                            else
                            {
                                _writer.WriteAttributeString("sz", "3600");
                            }
                            
                            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.GeometryTextBooleanProperties))
                            {
                                GeometryTextBooleanProperties gb = new GeometryTextBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.GeometryTextBooleanProperties].op);
                                if (gb.fUsegtextFKern & gb.gtextFKern)
                                {
                                    _writer.WriteAttributeString("kern", "10");
                                }
                            }

                            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
                            {
                                LineStyleBooleans lb = new LineStyleBooleans(so.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
                                if (lb.fUsefLine & lb.fLine)
                                {
                                    _writer.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", "000000");
                                    _writer.WriteEndElement();
                                    _writer.WriteEndElement();
                                    _writer.WriteEndElement();
                                }
                            }

                            _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", "FFFFFF");
                            _writer.WriteEndElement();
                            _writer.WriteEndElement();
                                                        
                            _writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("typeface", sFont);
                            _writer.WriteEndElement();
                            _writer.WriteEndElement();
                        }

                        _writer.WriteElementString("a", "t", OpenXmlNamespaces.DrawingML, sText);
                        _writer.WriteEndElement();
                    }

                    _writer.WriteElementString("a", "endParaRPr", OpenXmlNamespaces.DrawingML, "");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                }

                _writer.WriteEndElement();
            }
        }

        private void WriteShapeProperties(Shape sh, bool isPlaceholder, RegularContainer slide)
        {
            ShapeContainer container = (ShapeContainer)sh.ParentRecord;

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
            {
                FillStyleBooleanProperties p = new FillStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                if (p.fUsefFilled & p.fFilled) //  so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                {
                    new FillMapping(_ctx, _writer, parentSlideMapping).Apply(so);
                }
            }
            else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
            {
                if (sh.Instance != 0xca & isPlaceholder == false)
                {
                    string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);
                    _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", colorval);
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                    {
                        _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                        _writer.WriteEndElement();
                    }
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                }
            }

            _writer.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingML);
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineWidth))
            {
                _writer.WriteAttributeString("w", so.OptionsByID[ShapeOptions.PropertyId.lineWidth].op.ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndCapStyle))
            {
                switch (so.OptionsByID[ShapeOptions.PropertyId.lineEndCapStyle].op)
                {
                    case 0: //round
                        _writer.WriteAttributeString("cap", "rnd");
                        break;
                    case 1: //square
                        _writer.WriteAttributeString("cap", "sq");
                        break;
                    case 2: //flat
                        _writer.WriteAttributeString("cap", "flat");
                        break;
                }
            }


            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineType))
            {
                switch (so.OptionsByID[ShapeOptions.PropertyId.lineType].op)
                {
                    case 0: //solid
                        break;
                    case 1: //pattern
                        uint blipIndex = so.OptionsByID[ShapeOptions.PropertyId.lineFillBlip].op;
                        DrawingGroup gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                        BlipStoreEntry bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex - 1];
                        BitmapBlip b = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                        _writer.WriteStartElement("a", "pattFill", OpenXmlNamespaces.DrawingML);

                        _writer.WriteAttributeString("prst", Utils.getPrstForPatternCode(b.m_bTag)); //Utils.getPrstForPattern(blipNamePattern));

                        _writer.WriteStartElement("a", "fgClr", OpenXmlNamespaces.DrawingML);

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                        {
                            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, container.FirstAncestorWithType<Slide>(), so));
                            _writer.WriteEndElement();
                        }
                        else
                        {
                            _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", "tx1");
                            _writer.WriteEndElement();
                        }

                        _writer.WriteEndElement();

                        _writer.WriteStartElement("a", "bgClr", OpenXmlNamespaces.DrawingML);

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineBackColor))
                        {
                            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.lineBackColor].op, container.FirstAncestorWithType<Slide>(), so));
                            _writer.WriteEndElement();
                        }
                        else
                        {
                            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", "FFFFFF");
                            _writer.WriteEndElement();
                        }

                        _writer.WriteEndElement();

                        _writer.WriteEndElement();

                        break;
                    case 2: //texture
                        break;
                    default:
                        break;
                }
            }
            else
            {
                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor) & so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
                {
                    LineStyleBooleans lineStyle = new LineStyleBooleans(so.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
                    if (lineStyle.fLine)
                    {
                        string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, slide, so);
                        _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                        _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", colorval);
                        _writer.WriteEndElement();
                        _writer.WriteEndElement();
                    }
                }
            }

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineDashing))
            {
                _writer.WriteStartElement("a", "prstDash", OpenXmlNamespaces.DrawingML);
                switch ((ShapeOptions.LineDashing)so.OptionsByID[ShapeOptions.PropertyId.lineDashing].op)
                {
                    case ShapeOptions.LineDashing.Solid:
                        _writer.WriteAttributeString("val", "solid");
                        break;
                    case ShapeOptions.LineDashing.DashSys:
                        _writer.WriteAttributeString("val", "sysDash");
                        break;
                    case ShapeOptions.LineDashing.DotSys:
                        _writer.WriteAttributeString("val", "sysDot");
                        break;
                    case ShapeOptions.LineDashing.DashDotSys:
                        _writer.WriteAttributeString("val", "sysDashDot");
                        break;
                    case ShapeOptions.LineDashing.DashDotDotSys:
                        _writer.WriteAttributeString("val", "sysDashDotDot");
                        break;
                    case ShapeOptions.LineDashing.DotGEL:
                        _writer.WriteAttributeString("val", "dot");
                        break;
                    case ShapeOptions.LineDashing.DashGEL:
                        _writer.WriteAttributeString("val", "dash");
                        break;
                    case ShapeOptions.LineDashing.LongDashGEL:
                        _writer.WriteAttributeString("val", "lgDash");
                        break;
                    case ShapeOptions.LineDashing.DashDotGEL:
                        _writer.WriteAttributeString("val", "dashDot");
                        break;
                    case ShapeOptions.LineDashing.LongDashDotGEL:
                        _writer.WriteAttributeString("val", "lgDashDot");
                        break;
                    case ShapeOptions.LineDashing.LongDashDotDotGEL:
                        _writer.WriteAttributeString("val", "lgDashDotDot");
                        break;
                }
                _writer.WriteEndElement();
            }

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStartArrowhead))
            {
                ShapeOptions.LineEnd val = (ShapeOptions.LineEnd)so.OptionsByID[ShapeOptions.PropertyId.lineStartArrowhead].op;
                if (val != ShapeOptions.LineEnd.NoEnd)
                {
                    _writer.WriteStartElement("a", "headEnd", OpenXmlNamespaces.DrawingML);
                    switch (val)
                    {
                        case ShapeOptions.LineEnd.ArrowEnd:
                            _writer.WriteAttributeString("type", "triangle");
                            break;
                        case ShapeOptions.LineEnd.ArrowStealthEnd:
                            _writer.WriteAttributeString("type", "stealth");
                            break;
                        case ShapeOptions.LineEnd.ArrowDiamondEnd:
                            _writer.WriteAttributeString("type", "diamond");
                            break;
                        case ShapeOptions.LineEnd.ArrowOvalEnd:
                            _writer.WriteAttributeString("type", "oval");
                            break;
                        case ShapeOptions.LineEnd.ArrowOpenEnd:
                            _writer.WriteAttributeString("type", "arrow");
                            break;
                        case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                        case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                            _writer.WriteAttributeString("type", "triangle");
                            break;
                    }
                    _writer.WriteAttributeString("w", "med");
                    _writer.WriteAttributeString("len", "med");
                    _writer.WriteEndElement(); //headEnd
                }
            }

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndArrowhead))
            {
                ShapeOptions.LineEnd val = (ShapeOptions.LineEnd)so.OptionsByID[ShapeOptions.PropertyId.lineEndArrowhead].op;
                if (val != ShapeOptions.LineEnd.NoEnd)
                {
                    _writer.WriteStartElement("a", "tailEnd", OpenXmlNamespaces.DrawingML);
                    switch (val)
                    {
                        case ShapeOptions.LineEnd.ArrowEnd:
                            _writer.WriteAttributeString("type", "triangle");
                            break;
                        case ShapeOptions.LineEnd.ArrowStealthEnd:
                            _writer.WriteAttributeString("type", "stealth");
                            break;
                        case ShapeOptions.LineEnd.ArrowDiamondEnd:
                            _writer.WriteAttributeString("type", "diamond");
                            break;
                        case ShapeOptions.LineEnd.ArrowOvalEnd:
                            _writer.WriteAttributeString("type", "oval");
                            break;
                        case ShapeOptions.LineEnd.ArrowOpenEnd:
                            _writer.WriteAttributeString("type", "arrow");
                            break;
                        case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                        case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                            _writer.WriteAttributeString("type", "triangle");
                            break;
                    }
                    _writer.WriteAttributeString("w", "med");
                    _writer.WriteAttributeString("len", "med");
                    _writer.WriteEndElement(); //tailnd
                }
            }

            if (!so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndCapStyle))
            {
                //    _writer.WriteStartElement("a", "miter", OpenXmlNamespaces.DrawingML);
                //    _writer.WriteAttributeString("lim", "800000");
                //    _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); //ln

            //shadow
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ShadowStyleBooleanProperties))
            {
                ShadowStyleBooleanProperties sp = new ShadowStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.ShadowStyleBooleanProperties].op);
                if (sp.fUsefShadow & sp.fShadow)
                {
                    new ShadowMapping(_ctx, _writer).Apply(so);
                }
            }
        }

        private void writeOle(ShapeContainer container)
        {
            Shape sh = container.FirstChildWithType<Shape>();
            ShapeOptions so = container.FirstChildWithType<ShapeOptions>();

            _writer.WriteStartElement("p", "graphicFrame", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "nvGraphicFramePr", OpenXmlNamespaces.PresentationML);
            WriteCNvPr(--groupcounter, "");
            _writer.WriteStartElement("p", "cNvGraphicFramePr", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("a", "graphicFrameLocks", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("noChangeAspect", "1");
            _writer.WriteEndElement(); //graphicFrameLocks
            _writer.WriteEndElement(); //cNvGraphicFramePr
            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteEndElement(); //nvGraphicFramePr   

            ClientAnchor anchor = container.FirstChildWithType<ClientAnchor>();
            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                _writer.WriteStartElement("p", "xfrm", OpenXmlNamespaces.PresentationML);

                _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteEndElement();
            }

            _writer.WriteStartElement("a", "graphic", OpenXmlNamespaces.DrawingML);

            _writer.WriteStartElement("a", "graphicData", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("uri", OpenXmlRelationshipTypes.OleObject);


            OEPlaceHolderAtom placeholder = null;
            int exObjIdRef = 0;
            CheckClientData(container.FirstChildWithType<ClientData>(), ref placeholder, ref exObjIdRef);

            ExOleEmbedContainer oleContainer = this._ctx.Ppt.OleObjects[exObjIdRef];
           
            string spid = "";
            string rid = "";
            string width = "";
            string height = "";
            string name = "Chart";
            string progId = "MSGraph.Chart.8";

            _writer.WriteStartElement("p", "oleObj", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("spid", spid);
            _writer.WriteAttributeString("name", name);
            _writer.WriteAttributeString("id",OpenXmlNamespaces.Relationships, rid);
            _writer.WriteAttributeString("imgW", width);
            _writer.WriteAttributeString("imgH", height);
            _writer.WriteAttributeString("progId", progId);

            _writer.WriteStartElement("p", "embed", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("fllowColorScheme", "full");
            _writer.WriteEndElement(); //embed

            _writer.WriteEndElement(); //oleObj

            _writer.WriteEndElement(); //graphicData

            _writer.WriteEndElement(); //graphic


            _writer.WriteEndElement(); //graphicFrame
        }

        private void writePic(ShapeContainer container)
        {
            Shape sh = container.FirstChildWithType<Shape>();
            ShapeOptions so = container.FirstChildWithType<ShapeOptions>();

            uint indexOfPicture = 0;
            //TODO: read these infos from so
            ++_ctx.lastImageID;
            int id = _ctx.lastImageID; // "213";
            string name = "";
            string descr = "";
            string rId = "";
            foreach (ShapeOptions.OptionEntry en in so.Options)
            {

                switch (en.pid)
                {
                    case ShapeOptions.PropertyId.Pib:
                        indexOfPicture = en.op - 1;
                        break;
                    case ShapeOptions.PropertyId.pibName:
                    //    name = Encoding.Unicode.GetString(en.opComplex);
                    //    name = name.Substring(0, name.Length - 1).Replace("\0","");
                        break;
                    case ShapeOptions.PropertyId.pibPrintName:
                        id = (int)en.op;
                        break;

                }
            }

            DrawingGroup gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
            BlipStoreEntry bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)indexOfPicture];

            //if (this.parentSlideMapping is MasterMapping) return;
            
            if (this._ctx.AddedImages.ContainsKey(bse.foDelay))
            {
                rId = this._ctx.AddedImages[bse.foDelay]; 
            } else {

                if (!_ctx.Ppt.PicturesContainer._pictures.ContainsKey(bse.foDelay))
                {
                    return;
                }

                Record recBlip = _ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                if (recBlip is BitmapBlip)
                {
                    BitmapBlip b = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                    ImagePart imgPart = null;
                    imgPart = parentSlideMapping.targetPart.AddImagePart(getImageType(b.TypeCode));
                    imgPart.TargetDirectory = "..\\media";
                    System.IO.Stream outStream = imgPart.GetStream();
                    outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);

                    rId = imgPart.RelIdToString;
                }
                else if (recBlip is MetafilePictBlip)
                {
                    MetafilePictBlip mb = (MetafilePictBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                    ImagePart imgPart = null;
                    imgPart = parentSlideMapping.targetPart.AddImagePart(getImageType(mb.TypeCode));
                    imgPart.TargetDirectory = "..\\media";
                    System.IO.Stream outStream = imgPart.GetStream();

                    byte[] decompressed = mb.Decrompress();
                    outStream.Write(decompressed, 0, decompressed.Length);
                    //outStream.Write(mb.m_pvBits, 0, mb.m_pvBits.Length);

                    rId = imgPart.RelIdToString;
                }
                //this._ctx.AddedImages.Add(bse.foDelay, rId);
            }

            _writer.WriteStartElement("p", "pic", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "nvPicPr", OpenXmlNamespaces.PresentationML);


            if (!spidToId.ContainsKey(sh.spid))
            {
                spidToId.Add(sh.spid, id);
            }

            _writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", spidToId[sh.spid].ToString());

            //_writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            //_writer.WriteAttributeString("id", id);
            _writer.WriteAttributeString("name", name);
            _writer.WriteAttributeString("descr", descr);
            _writer.WriteEndElement(); //p:cNvPr

            _writer.WriteStartElement("p", "cNvPicPr", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("a", "picLocks", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("noChangeAspect", "1");
            _writer.WriteAttributeString("noChangeArrowheads", "1");
            _writer.WriteEndElement(); //a:picLocks
            _writer.WriteEndElement(); //p:cNvPicPr

            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");

            _writer.WriteEndElement(); //p:nvPicPr

            _writer.WriteStartElement("p", "blipFill", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("embed", OpenXmlNamespaces.Relationships, rId);

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureBrightness) | so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureContrast))
            {
                _writer.WriteStartElement("a", "lum", OpenXmlNamespaces.DrawingML);

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureBrightness))
                {
                    UInt32 b = so.OptionsByID[ShapeOptions.PropertyId.pictureBrightness].op;
                    if (b == 0xFFF8000) b = 0;

                    Decimal b1 = 0;

                    if (((int)b) < 0)
                    {
                        int b2 = (int)b;
                        b1 = (Decimal)b2 / 0x8000;
                    }
                    else
                    {
                        b1 = (Decimal)b / 0x8000;
                    }
                    
                    b1 = b1 * 100000;
                    b1 = Math.Floor(b1);
                    _writer.WriteAttributeString("bright", b1.ToString());
                }

                //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureContrast))
                //{
                //    Int32 b = (int)so.OptionsByID[ShapeOptions.PropertyId.pictureContrast].op;
                //    Decimal b2 = (Decimal)b  / 0x10000; //This comes from analysing, no hint in spec found
                //    b2 = b2 * 100000;
                //    b2 = Math.Floor(b2);
                //    if (b == 0x7FFFFFFF) b2 = 100000;
                //    if (b2 > 100000) b2 = 100000;
                    
                //    _writer.WriteAttributeString("contrast", b2.ToString());
                //}

                _writer.WriteEndElement();
            }
           

            _writer.WriteEndElement(); //a:blip
            _writer.WriteStartElement("a", "srcRect", OpenXmlNamespaces.DrawingML);
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromLeft))
            {
                _writer.WriteAttributeString("l", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromLeft].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromTop))
            {
                _writer.WriteAttributeString("t", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromTop].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromRight))
            {
                _writer.WriteAttributeString("r", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromRight].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromBottom))
            {
                _writer.WriteAttributeString("b", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromBottom].op / 65536 * 100000).ToString());
            }
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingML);
            _writer.WriteElementString("a", "fillRect", OpenXmlNamespaces.DrawingML, "");
            _writer.WriteEndElement(); //a:stretch
            _writer.WriteEndElement(); //p:blipFill

            // Visible shape properties
            _writer.WriteStartElement("p", "spPr", OpenXmlNamespaces.PresentationML);
            ClientAnchor anchor = container.FirstChildWithType<ClientAnchor>();
            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

                _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteEndElement();
            }

            _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("prst", "rect");
            _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
            _writer.WriteEndElement(); //a:prstGeom
            _writer.WriteElementString("a", "noFill", OpenXmlNamespaces.DrawingML, "");

            //shadow
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ShadowStyleBooleanProperties))
            {
                ShadowStyleBooleanProperties p = new ShadowStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.ShadowStyleBooleanProperties].op);
                if (p.fUsefShadow & p.fShadow)
                {
                    new ShadowMapping(_ctx, _writer).Apply(so);
                }
            }

            _writer.WriteEndElement(); //p:spPr
            
            _writer.WriteEndElement(); //p:pic

        }

        public void writeBodyPr(Record rec)
        {
            _writer.WriteStartElement("a", "bodyPr", OpenXmlNamespaces.DrawingML);
            bool cancelAttributes = false;
            if (rec is ShapeContainer)
            {
                ShapeContainer container = (ShapeContainer)rec;
                switch (container.FirstChildWithType<Shape>().Instance)
                {
                    case 0x88:
                        _writer.WriteAttributeString("wrap", "none");

                        _writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "textPlain");
                        _writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    default:
                        break;
                }
            }

            if (!cancelAttributes)
            {

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextLeft))
                {
                    _writer.WriteAttributeString("lIns", so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op.ToString());
                }

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dyTextTop))
                {
                    _writer.WriteAttributeString("tIns", so.OptionsByID[ShapeOptions.PropertyId.dyTextTop].op.ToString());
                }

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextRight))
                {
                    _writer.WriteAttributeString("rIns", so.OptionsByID[ShapeOptions.PropertyId.dxTextRight].op.ToString());
                }

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dyTextBottom))
                {
                    _writer.WriteAttributeString("bIns", so.OptionsByID[ShapeOptions.PropertyId.dyTextBottom].op.ToString());
                }


                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.WrapText))
                {
                    switch (so.OptionsByID[ShapeOptions.PropertyId.WrapText].op)
                    {
                        case 0: //square
                            _writer.WriteAttributeString("wrap", "square");
                            break;
                        case 1: //by points
                            break; //TODO
                        case 2: //none
                            _writer.WriteAttributeString("wrap", "none");
                            break;
                        case 3: //top bottom
                        case 4: //through
                        default:
                            break; //TODO
                    }
                }

                string s = "";
                foreach (ShapeOptions.OptionEntry en in so.Options)
                {
                    switch (en.pid)
                    {
                        case ShapeOptions.PropertyId.anchorText:

                            switch (en.op)
                            {
                                case 0: //Top
                                    _writer.WriteAttributeString("anchor", "t");
                                    break;
                                case 1: //Middle
                                    _writer.WriteAttributeString("anchor", "ctr");
                                    break;
                                case 2: //Bottom
                                    _writer.WriteAttributeString("anchor", "b");
                                    break;
                                case 3: //TopCentered
                                    _writer.WriteAttributeString("anchor", "t");
                                    _writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 4: //MiddleCentered
                                    _writer.WriteAttributeString("anchor", "ctr");
                                    _writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 5: //BottomCentered
                                    _writer.WriteAttributeString("anchor", "b");
                                    _writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 6: //TopBaseline
                                case 7: //BottomBaseline
                                case 8: //TopCenteredBaseline
                                case 9: //BottomCenteredBaseline
                                    //TODO
                                    break;
                            }
                            break;
                        default:
                            s += en.pid.ToString() + " ";
                            break;
                    }
                }
            }

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.TextBooleanProperties))
            {
                TextBooleanProperties props = new TextBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.TextBooleanProperties].op);
                if (props.fFitShapeToText && props.fUsefFitShapeToText) _writer.WriteElementString("a", "spAutoFit", OpenXmlNamespaces.DrawingML, "");
            }

            _writer.WriteEndElement();
        }

        private void CheckClientData(ClientData clientData, ref OEPlaceHolderAtom placeholder, ref int exObjIdRef)
        {
            ShapeStyleTextProp9Atom = null;
            bool phWritten = false;
            if (clientData != null)
            {

                System.IO.MemoryStream ms = new System.IO.MemoryStream(clientData.bytes);

                if (ms.Length > 0)
                {

                    Record rec = Record.ReadRecord(ms, 0);

                    if (rec.TypeCode == 4116)
                    {
                        AnimationInfoContainer animinfo = (AnimationInfoContainer)rec;
                        animinfos.Add(animinfo, _idCnt);
                        if (ms.Position < ms.Length) rec = Record.ReadRecord(ms, 1);
                    }

                    switch (rec.TypeCode)
                    {
                        case 3009:
                            exObjIdRef = ((ExObjRefAtom)rec).exObjIdRef;
                            break;
                        case 3011:
                            placeholder = (OEPlaceHolderAtom)rec;

                            if (placeholder != null)
                            {

                                _writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);

                                if (!placeholder.IsObjectPlaceholder())
                                {
                                    string typeValue = Utils.PlaceholderIdToXMLValue(placeholder.PlacementId);
                                    _writer.WriteAttributeString("type", typeValue);
                                }

                                switch (placeholder.PlaceholderSize)
                                {
                                    case 1:
                                        _writer.WriteAttributeString("sz", "half");
                                        break;
                                    case 2:
                                        _writer.WriteAttributeString("sz", "quarter");
                                        break;
                                }


                                if (placeholder.Position != -1)
                                {
                                    _writer.WriteAttributeString("idx", placeholder.Position.ToString());
                                }
                                else
                                {
                                    try
                                    {
                                        Slide master = _ctx.Ppt.FindMasterRecordById(clientData.FirstAncestorWithType<Slide>().FirstChildWithType<SlideAtom>().MasterId);
                                        foreach (ShapeContainer cont in master.FirstChildWithType<PPDrawing>().FirstChildWithType<DrawingContainer>().FirstChildWithType<GroupContainer>().AllChildrenWithType<ShapeContainer>())
                                        {
                                            Shape s = cont.FirstChildWithType<Shape>();
                                            ClientData d = cont.FirstChildWithType<ClientData>();
                                            if (d != null)
                                            {
                                                ms = new System.IO.MemoryStream(d.bytes);
                                                rec = Record.ReadRecord(ms, 0);
                                                if (rec is OEPlaceHolderAtom)
                                                {
                                                    OEPlaceHolderAtom placeholder2 = (OEPlaceHolderAtom)rec;
                                                    if (placeholder2.PlacementId == PlaceholderEnum.MasterBody && (placeholder.PlacementId == PlaceholderEnum.Body || placeholder.PlacementId == PlaceholderEnum.Object))
                                                    {
                                                        if (placeholder2.Position != -1)
                                                        {
                                                            _writer.WriteAttributeString("idx", placeholder2.Position.ToString());
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                    }

                                    catch (Exception)
                                    {
                                       //ignore
                                    }
                                }
                                
                                _writer.WriteEndElement();
                                phWritten = true;
                            }
                            break;
                        case 4116:
                            break;
                        case 5000:
                            RegularContainer con = (RegularContainer)rec;
                            foreach (ProgBinaryTag t in con.AllChildrenWithType<ProgBinaryTag>())
                            {
                                CStringAtom c = t.FirstChildWithType<CStringAtom>();
                                ProgBinaryTagDataBlob b = t.FirstChildWithType<ProgBinaryTagDataBlob>();
                                StyleTextProp9Atom p = b.FirstChildWithType<StyleTextProp9Atom>();
                                ShapeStyleTextProp9Atom = p;
                            }
                            break;
                        default:
                            break;
                    }
                }
                               
            
                RegularContainer container = (RegularContainer)(clientData.ParentRecord);
                foreach (ClientTextbox b in container.AllChildrenWithType<ClientTextbox>())
                {
                    ms = new System.IO.MemoryStream(b.Bytes);
                    Record rec;
                    while (ms.Position < ms.Length)
                    {
                        rec = Record.ReadRecord(ms, 0);

                        switch (rec.TypeCode)
                        {
                            case 0xfa0: //TextCharsAtom
                            case 0xfa1: //TextRunStyleAtom
                            case 0xfa6: //TextRulerAtom
                            case 0xfa8: //TextBytesAtom
                            case 0xfaa: //TextSpecialInfoAtom
                            case 0xfa2: //MasterTextPropAtom
                                break;
                            case 0xfd8: //SlideNumberMCAtom
                               
                                break;
                            case 0xff7: //DateTimeMCAtom
                                if (!phWritten)
                                {
                                    _writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);
                                    _writer.WriteAttributeString("type", "dt");
                                    _writer.WriteEndElement();
                                }
                                break;
                            case 0xff9: //HeaderMCAtom
                                break;
                            case 0xffa: //FooterMCAtom
                                break;
                            case 0xff8: //GenericDateMCAtom
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
        }

        public void Apply(ClientTextbox textbox)
        {
            Apply(textbox, false);
        }

        public void Apply(ClientTextbox textbox, bool insideTable)
        {
            if (insideTable)
            {
                _writer.WriteStartElement("a", "txBody", OpenXmlNamespaces.DrawingML);
            }
            else
            {
                _writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);
            }

            writeBodyPr(textbox);

            _writer.WriteStartElement("a", "lstStyle", OpenXmlNamespaces.DrawingML);
            bool lvlRprWritten = false;

            System.IO.MemoryStream ms = new System.IO.MemoryStream(textbox.Bytes);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            List<int> lst = new List<int>();
            string lang = "en-US";
            while (ms.Position < ms.Length)
            {
                Record rec = Record.ReadRecord(ms, 0);

                switch (rec.TypeCode)
                {
                    case 0xf9e: //OutlineTextRefAtom
                        OutlineTextRefAtom otrAtom = (OutlineTextRefAtom)rec;
                        SlideListWithText slideListWithText = _ctx.Ppt.DocumentRecord.RegularSlideListWithText;

                        List<TextHeaderAtom> thAtoms = slideListWithText.SlideToPlaceholderTextHeaders[textbox.FirstAncestorWithType<Slide>().PersistAtom];
                        thAtom = thAtoms[otrAtom.Index];

                        //if (thAtom.TextAtom != null) text = thAtom.TextAtom.Text;
                        if (thAtom.TextStyleAtom != null) style = thAtom.TextStyleAtom;

                        break;
                    case 0xf9f: //TextHeaderAtom
                        thAtom = (TextHeaderAtom)rec;
                        break;
                    case 0xfa0: //TextCharsAtom
                        thAtom.TextAtom = (TextAtom)rec;
                        break;
                    case 0xfa1: //StyleTextPropAtom
                        style = (TextStyleAtom)rec;
                        style.TextHeaderAtom = thAtom;
                        break;
                    case 0xfa2: //MasterTextPropAtom
                        MasterTextPropAtom m = (MasterTextPropAtom)rec;
                        foreach(MasterTextPropRun r in m.MasterTextPropRuns)
                        {
                            if (!lst.Contains(r.indentLevel))
                            {
                                _writer.WriteStartElement("a", "lvl" + (r.indentLevel + 1) + "pPr", OpenXmlNamespaces.DrawingML);

                                if (thAtom.TextType == TextType.CenterTitle || thAtom.TextType == TextType.CenterBody)
                                {
                                    _writer.WriteAttributeString("algn", "ctr");
                                }

                                //_writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");

                                _writer.WriteEndElement();
                                lst.Add(r.indentLevel);
                            }
                        }
                        break;
                    case 0xfa8: //TextBytesAtom
                        //text = ((TextBytesAtom)rec).Text;
                        thAtom.TextAtom = (TextAtom)rec;
                        break;
                    case 0xfaa: //TextSpecialInfoAtom
                        TextSpecialInfoAtom sia = (TextSpecialInfoAtom)rec;
                        if (sia.Runs.Count > 0)
                        {
                            if (sia.Runs[0].si.lang)
                            {
                                switch (sia.Runs[0].si.lid)
                                {
                                    case 0x0: // no language
                                        break;
                                    case 0x13: //Any Dutch language is preferred over non-Dutch languages when proofing the text
                                        break;
                                    case 0x400: //no proofing
                                        break;
                                    default:
                                        try
                                        {
                                            lang = System.Globalization.CultureInfo.GetCultureInfo(sia.Runs[0].si.lid).IetfLanguageTag;
                                        }
                                        catch (Exception)
                                        {   
                                            //ignore
                                        }                                       
                                        break;
                                }                               
                            }
                        }
                        break;
                    case 0xfd8: //SlideNumberMCAtom
                    case 0xff9: //HeaderMCAtom
                    case 0xffa: //FooterMCAtom
                    case 0xff8: //GenericDateMCAtom
                        if (!lvlRprWritten)
                        foreach (ParagraphRun r in style.PRuns)
                        {
                            _writer.WriteStartElement("a", "lvl" + (r.IndentLevel + 1) + "pPr", OpenXmlNamespaces.DrawingML);
                            if (r.AlignmentPresent)
                            {
                                switch (r.Alignment)
                                {
                                    case 0x0000: //Left
                                        _writer.WriteAttributeString("algn", "l");
                                        break;
                                    case 0x0001: //Center
                                        _writer.WriteAttributeString("algn", "ctr");
                                        break;
                                    case 0x0002: //Right
                                        _writer.WriteAttributeString("algn", "r");
                                        break;
                                    case 0x0003: //Justify
                                        _writer.WriteAttributeString("algn", "just");
                                        break;
                                    case 0x0004: //Distributed
                                        _writer.WriteAttributeString("algn", "dist");
                                        break;
                                    case 0x0005: //ThaiDistributed
                                        _writer.WriteAttributeString("algn", "thaiDist");
                                        break;
                                    case 0x0006: //JustifyLow
                                        _writer.WriteAttributeString("algn", "justLow");
                                        break;
                                }
                            }
                            string lastColor = "";
                            string lastSize = "";
                            string lastTypeface = "";
                            RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                            new CharacterRunPropsMapping(_ctx, _writer).Apply(style.CRuns[0], "defRPr", slide, ref lastColor, ref lastSize, ref lastTypeface, lang, null); 
                            _writer.WriteEndElement();
                            lvlRprWritten = true;
                        }
                        break;
                    default:
                        break;
                }
            }

            _writer.WriteEndElement();

            new TextMapping(_ctx, _writer).Apply(this, textbox, _footertext, _headertext, _datetext);

            _writer.WriteEndElement();
        }

        public static ImagePart.ImageType getImageType(uint TypeCode)
        {
            switch (TypeCode)
            {
                case 0xF01A:
                    return ImagePart.ImageType.Emf;
                case 0xF01B:
                    return ImagePart.ImageType.Wmf;
                case 0xF01D:
                    return ImagePart.ImageType.Jpeg;
                case 0xF01E:
                    return ImagePart.ImageType.Png;
                case 0xF01F: //DIP
                    return ImagePart.ImageType.Bmp;
                case 0xF020:
                    return ImagePart.ImageType.Tiff;
                default:
                    return ImagePart.ImageType.Png;
            }
        }


        public void Apply(RegularContainer container)
        {
            // Descend into container records by default
            foreach (Record record in container.Children)
            {
                DynamicApply(record);
            }
        }

        public void Apply(Record record)
        {
            // Ignore unsupported records
            //TraceLogger.DebugInternal("Unsupported record: {0}", record);
        }

        private void WriteGroupShapeProperties(ShapeContainer header)
        {
            GroupShapeRecord groupShape = header.FirstChildWithType<GroupShapeRecord>();

            // Write non-visible Group Shape properties
            _writer.WriteStartElement("p", "nvGrpSpPr", OpenXmlNamespaces.PresentationML);

            // Non-visible Canvas Properties
            WriteCNvPr(-1, "");

            _writer.WriteElementString("p", "cNvGrpSpPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");

            _writer.WriteEndElement();


            // Write visible Group Shape properties
            _writer.WriteStartElement("p", "grpSpPr", OpenXmlNamespaces.PresentationML);
            WriteXFrm(_writer, new Rectangle()); // groupShape.rcgBounds

            _writer.WriteEndElement();
        }

        private void WritecustGeom(Shape sh)
        {

            if (!so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pVertices) | !so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pSegmentInfo))
            {
                _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("prst", "rect");
                _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                _writer.WriteEndElement(); //prstGeom
                return;
            }

            _writer.WriteStartElement("a", "custGeom", OpenXmlNamespaces.DrawingML);

            _writer.WriteStartElement("a", "cxnLst", OpenXmlNamespaces.DrawingML);

            ShapeOptions.OptionEntry pVertices = so.OptionsByID[ShapeOptions.PropertyId.pVertices];
            
            uint shapepath = 1;
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shapePath))
            {
                ShapeOptions.OptionEntry ShapePath = so.OptionsByID[ShapeOptions.PropertyId.shapePath];
                shapepath = ShapePath.op;
            }
            ShapeOptions.OptionEntry SegementInfo = so.OptionsByID[ShapeOptions.PropertyId.pSegmentInfo];
            PathParser pp = new PathParser(SegementInfo.opComplex, pVertices.opComplex);

            foreach (Point point in pp.Values)
            {
                _writer.WriteStartElement("a", "cxn", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("ang", "0");
                _writer.WriteStartElement("a", "pos", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("x", point.X.ToString());
                _writer.WriteAttributeString("y", point.Y.ToString());
                _writer.WriteEndElement(); //pos
                _writer.WriteEndElement(); //cxn
            }
            _writer.WriteEndElement(); //cxnLst

            _writer.WriteStartElement("a", "rect", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("l", "0");
            _writer.WriteAttributeString("t", "0");
            _writer.WriteAttributeString("r", "r");
            _writer.WriteAttributeString("b", "b");
            _writer.WriteEndElement(); //rect

            _writer.WriteStartElement("a", "pathLst", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "path", OpenXmlNamespaces.DrawingML);
            //compute width and height:
            int minX = 1000;
            int minY = 1000;
            int maxX = -1000;
            int maxY = -1000;
            foreach (Point p in pp.Values)
            {
                if ((p.X) < minX) minX = p.X;
                if ((p.X) > maxX) maxX = p.X;
                if ((p.Y) < minY) minY = p.Y;
                if ((p.Y) > maxY) maxY = p.Y;
            }
            _writer.WriteAttributeString("w", (maxX - minX).ToString());
            _writer.WriteAttributeString("h", (maxY - minY).ToString());

            int valuePointer = 0;

            switch (shapepath)
            {
                case 0: //lines
                case 1: //lines closed
                    while (valuePointer < pp.Values.Count)
                    {
                        if (valuePointer == 0)
                        {
                            _writer.WriteStartElement("a", "moveTo", OpenXmlNamespaces.DrawingML);
                            _writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                            _writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                            _writer.WriteEndElement(); //pr
                            _writer.WriteEndElement(); //moveTo
                            valuePointer += 1;
                        }
                        else
                        {
                            _writer.WriteStartElement("a", "lnTo", OpenXmlNamespaces.DrawingML);
                            _writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                            _writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                            _writer.WriteEndElement(); //pt
                            _writer.WriteEndElement(); //lnTo
                            valuePointer += 1;
                        }
                    }
                    break;
                case 2: //curves
                    break;
                case 3: //curves closed
                    break;
                case 4: //complex
                    foreach (PathSegment seg in pp.Segments)
                    {
                        if (valuePointer >= pp.Values.Count) break;
                        switch (seg.Type)
                        {
                            case PathSegment.SegmentType.msopathLineTo:
                                _writer.WriteStartElement("a", "lnTo", OpenXmlNamespaces.DrawingML);
                                _writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                _writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                _writer.WriteEndElement(); //pt
                                _writer.WriteEndElement(); //lnTo
                                valuePointer += 1;
                                break;
                            case PathSegment.SegmentType.msopathCurveTo:
                                _writer.WriteStartElement("a", "cubicBezTo", OpenXmlNamespaces.DrawingML);
                                _writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                _writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                _writer.WriteEndElement(); //pt
                                valuePointer += 1;
                                _writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                _writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                _writer.WriteEndElement(); //pt
                                valuePointer += 1;
                                _writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                _writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                _writer.WriteEndElement(); //pt
                                valuePointer += 1;
                                _writer.WriteEndElement(); //cubicBezTo
                                break;
                            case PathSegment.SegmentType.msopathMoveTo:
                                _writer.WriteStartElement("a", "moveTo", OpenXmlNamespaces.DrawingML);
                                _writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                _writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                _writer.WriteEndElement(); //pr
                                _writer.WriteEndElement(); //moveTo
                                valuePointer += 1;
                                break;
                            case PathSegment.SegmentType.msopathClose:
                                _writer.WriteElementString("a", "close", OpenXmlNamespaces.DrawingML, "");
                                break;
                            default:
                                break;
                        }
                    }
                    break;
            }

            _writer.WriteEndElement(); //path
            _writer.WriteEndElement(); //pathLst

            _writer.WriteEndElement(); //custGeom
        }

        private void WriteprstGeom(Shape shape)
        {
            if (shape != null)
            {
                string prst = Utils.getPrstForShape(shape.Instance);
                if (prst.Length > 0)
                {
                   
                    _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("prst", prst);
                    if (prst == "roundRect" & so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.adjustValue)) //TODO: implement for all shapes
                    {
                        _writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        _writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("name", "adj");
                        _writer.WriteAttributeString("fmla", "val " + Math.Floor(so.OptionsByID[ShapeOptions.PropertyId.adjustValue].op * 4.63).ToString()); //TODO: find out where this 4.63 comes from (value found by analysing behaviour of Powerpoint 2003)
                        _writer.WriteEndElement();
                        _writer.WriteEndElement();
                    }
                    else if (prst == "wedgeRectCallout" & so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.adjustValue))
                    {
                        //the following computations are based on experiments using Powerpoint 2003 and are not part of the spec
                        Decimal val = (Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.adjustValue].op;
                        Decimal percent = val / 21600 * 100;
                        int newVal = 0;
                        if (percent >= 50)
                        {
                            newVal = (int)(percent - 50) * 1000;
                        }
                        else
                        {
                            newVal = (int)(50 - percent) * -1000;
                        }

                        _writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        _writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("name", "adj1");
                        _writer.WriteAttributeString("fmla", "val " + newVal.ToString()); 
                        _writer.WriteEndElement();
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.adjust2Value))
                        {
                            val = (Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.adjust2Value].op;
                            percent = val / 21600 * 100;
                            newVal = 0;
                            if (percent >= 50)
                            {
                                newVal = (int)(percent - 50) * 1000;
                            }
                            else
                            {
                                newVal = (int)(50 - percent) * -1000;
                            }
                            _writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("name", "adj2");
                            _writer.WriteAttributeString("fmla", "val " + newVal.ToString()); 
                            _writer.WriteEndElement();
                        }
                        _writer.WriteEndElement();                       

                    } else {
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                    }
                    _writer.WriteEndElement(); //prstGeom
                }                
            }
        }

        

        public Dictionary<int, int> spidToId = new Dictionary<int, int>();
        private void WriteCNvPr(int spid, string name)
        {

            if (!spidToId.ContainsKey(spid))
            {
                spidToId.Add(spid, ++_idCnt);
            }

            _writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", spidToId[spid].ToString());
            _writer.WriteAttributeString("name", name);
            _writer.WriteEndElement();
            //if (!spidToId.ContainsKey(spid)) spidToId.Add(spid, _idCnt);
        }

        private void WriteXFrm(XmlWriter _writer, Rectangle rect)
        {
            _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

            // TODO: Coordinate conversion?
            _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("x", rect.X.ToString());
            _writer.WriteAttributeString("y", rect.Y.ToString());
            _writer.WriteEndElement();

            // TODO: Coordinate conversion?
            _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("cx", rect.Width.ToString());
            _writer.WriteAttributeString("cy", rect.Height.ToString());
            _writer.WriteEndElement();

            // TODO: Where do we get this from?
            _writer.WriteStartElement("a", "chOff", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("x", "0");
            _writer.WriteAttributeString("y", "0");
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "chExt", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("cx", rect.Width.ToString());
            _writer.WriteAttributeString("cy", rect.Height.ToString());
            _writer.WriteEndElement();            

            _writer.WriteEndElement();
        }
    }
}
