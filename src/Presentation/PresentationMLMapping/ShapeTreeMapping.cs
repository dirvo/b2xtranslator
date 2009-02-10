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
    class ShapeTreeMapping :
        AbstractOpenXmlMapping,
        IMapping<PPDrawing>
    {
        protected int _idCnt;
        protected ConversionContext _ctx;

        public PresentationMapping<Slide> parentSlideMapping = null;
        public Dictionary<AnimationInfoContainer, int> animinfos = new Dictionary<AnimationInfoContainer, int>();

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

        private ShapeOptions so = null;
        public void Apply(ShapeContainer container)
        {
            ClientData clientData = container.FirstChildWithType<ClientData>();

            bool pictureWritten = false;

            Shape sh = container.FirstChildWithType<Shape>();
            so = container.FirstChildWithType<ShapeOptions>();
            if (clientData == null)
            if (so != null)
            {
                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.Pib))
                {
                    writePic(container);
                    pictureWritten = true;
                }

                //foreach (ShapeOptions.OptionEntry en in so.Options)
                //{
                //    if (en.pid == ShapeOptions.PropertyId.Pib)
                //    {
                //        if (en.opComplex != null)
                //        {
                //            //TODO
                //        }
                //        else
                //        {
                //            writePic(container);
                //            pictureWritten = true;
                //        }
                //    }
                //}
            }

            if (!pictureWritten)
            {
                //bool isConnector = false;
                //switch (sh.Instance)
                //{
                //    case 0x20:
                //    case 0x21:
                //    case 0x22:
                //    case 0x23:
                //    case 0x24:
                //    case 0x25:
                //    case 0x26:
                //    case 0x27:
                //    case 0x28:
                //        isConnector = true;
                //        break;
                //    default:
                //        break;
                //}

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



                if (clientData != null)
                {

                    System.IO.MemoryStream ms = new System.IO.MemoryStream(clientData.bytes);
                    Record rec = Record.ReadRecord(ms, 0);

                    if (rec.TypeCode == 4116)
                    {
                        AnimationInfoContainer animinfo = (AnimationInfoContainer)rec;
                        animinfos.Add(animinfo, _idCnt);
                        if (ms.Position < ms.Length) rec = Record.ReadRecord(ms, 1);
                    }

                    switch (rec.TypeCode)
                    {
                        case 3011:
                            OEPlaceHolderAtom placeholder = (OEPlaceHolderAtom)rec; // clientData.FirstChildWithType<OEPlaceHolderAtom>();

                            if (placeholder != null)
                            {

                                _writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);

                                if (!placeholder.IsObjectPlaceholder())
                                {
                                    string typeValue = Utils.PlaceholderIdToXMLValue(placeholder.PlacementId);
                                    _writer.WriteAttributeString("type", typeValue);
                                }

                                if (placeholder.Position != -1)
                                {
                                    _writer.WriteAttributeString("idx", placeholder.Position.ToString());
                                }

                                _writer.WriteEndElement();
                            }
                            break;
                        case 4116:
                            //AnimationInfoContainer animinfo = (AnimationInfoContainer)rec;
                            //animinfos.Add(animinfo, _idCnt);
                            //new AnimationMapping(_ctx, _writer).Apply(animinfo);
                            break;
                    }
                }

                _writer.WriteEndElement();

                _writer.WriteEndElement();


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
                if (sh.Instance != 0)
                {
                    WriteprstGeom(sh);
                }
                else
                {
                    _writer.WriteStartElement("a", "custGeom", OpenXmlNamespaces.DrawingML);

                    _writer.WriteStartElement("a", "cxnLst", OpenXmlNamespaces.DrawingML);
                    
                    ShapeOptions.OptionEntry pVertices = so.OptionsByID[ShapeOptions.PropertyId.pVertices];
                    ShapeOptions.OptionEntry ShapePath = so.OptionsByID[ShapeOptions.PropertyId.shapePath];
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
                    foreach (PathSegment seg in pp.Segments)
                    {
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
                        }
                    }

                    _writer.WriteEndElement(); //path
                    _writer.WriteEndElement(); //pathLst
                    
                    _writer.WriteEndElement(); //custGeom
                }

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                {
                    string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, container.FirstAncestorWithType<Slide>());
                    _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", colorval);

                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                    {
                        _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString());
                        _writer.WriteEndElement();
                    }

                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                }

                _writer.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingML);
                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineWidth))
                {
                    _writer.WriteAttributeString("w", so.OptionsByID[ShapeOptions.PropertyId.lineWidth].op.ToString());
                }
                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndCapStyle))
                {
                    switch(so.OptionsByID[ShapeOptions.PropertyId.lineEndCapStyle].op)
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

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                {
                    string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, container.FirstAncestorWithType<Slide>());
                    _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", colorval);
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
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
                _writer.WriteEndElement(); //ln

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
                    //write dummy
                    _writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);
                    _writer.WriteElementString("a", "bodyPr", OpenXmlNamespaces.DrawingML,"");
                    _writer.WriteElementString("a", "lstStyle",OpenXmlNamespaces.DrawingML,"");
                    _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);
                    _writer.WriteElementString("a", "endParaRPr", OpenXmlNamespaces.DrawingML,"");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                }

                _writer.WriteEndElement();
            }
        }

        private void writePic(ShapeContainer container)
        {
            Shape sh = container.FirstChildWithType<Shape>();
            ShapeOptions so = container.FirstChildWithType<ShapeOptions>();

            uint indexOfPicture = 0;
            //TODO: read these infos from so
            ++_ctx.lastImageID;
            string id = _ctx.lastImageID.ToString(); // "213";
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
                        name = Encoding.Unicode.GetString(en.opComplex);
                        name = name.Substring(0, name.Length - 1);
                        break;
                    case ShapeOptions.PropertyId.pibPrintName:
                        id = en.op.ToString();
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

                BitmapBlip b = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                ImagePart imgPart = null;
                imgPart = parentSlideMapping.targetPart.AddImagePart(getImageType(b.TypeCode));
                imgPart.TargetDirectory = "..\\media";
                System.IO.Stream outStream = imgPart.GetStream();
                outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);

                rId = imgPart.RelIdToString;
                //this._ctx.AddedImages.Add(bse.foDelay, rId);
            }

            _writer.WriteStartElement("p", "pic", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "nvPicPr", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", id);
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
            _writer.WriteEndElement(); //a:blip
            _writer.WriteElementString("a", "srcRect", OpenXmlNamespaces.DrawingML, "");
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

            _writer.WriteEndElement(); //p:spPr

            

            _writer.WriteEndElement(); //p:pic
        }

        public void Apply(ClientTextbox textbox)
        {
            _writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("a", "bodyPr", OpenXmlNamespaces.DrawingML);

            string s = "";
            foreach (ShapeOptions.OptionEntry en in so.Options)
            {
                switch(en.pid)
                {
                    case ShapeOptions.PropertyId.anchorText:
                        
                        switch (en.op)
                        {
                            case 0: //Top
                                _writer.WriteAttributeString("anchor", "t");
                                break;
                            case 1: //Middle
                                _writer.WriteAttributeString("anchor","ctr");
                                break;
                            case 2: //Bottom
                                _writer.WriteAttributeString("anchor", "b");
                                break;
                            case 3: //TopCentered
                            case 4: //MiddleCentered
                            case 5: //BottomCentered
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
            
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "lstStyle", OpenXmlNamespaces.DrawingML);

            System.IO.MemoryStream ms = new System.IO.MemoryStream(textbox.Bytes);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
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

                            _writer.WriteStartElement("a", "lvl" + (r.indentLevel + 1) + "pPr", OpenXmlNamespaces.DrawingML);

                            if (thAtom.TextType == TextType.CenterTitle || thAtom.TextType == TextType.CenterBody)
                            {
                                _writer.WriteAttributeString("algn", "ctr");                                                                
                            }

                            //_writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");

                            _writer.WriteEndElement();
                            
                        }
                        break;
                    case 0xfa8: //TextBytesAtom
                        //text = ((TextBytesAtom)rec).Text;
                        thAtom.TextAtom = (TextAtom)rec;
                        break;
                    case 0xfaa: //TextSpecialInfoAtom
                        break;
                    case 0xfd8: //SlideNumberMCAtom
                        break;
                    case 0xff8: //GenericDateMCAtom
                        break;
                    default:
                        break;
                }
            }

            _writer.WriteEndElement();

            new TextMapping(_ctx, _writer).Apply(textbox);

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

        private void WriteprstGeom(Shape shape)
        {
            if (shape != null)
            {
                switch (shape.Instance)
                {
                    case 0x0: //NotPrimitive
                    case 0x1: //Rectangle
                        break;
                    case 0x2: //RoundRectangle
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "roundRect");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x3: //ellipse
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst","ellipse");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x4: //diamond
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "diamond");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x5: //triangle
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "triangle");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x6: //right triangle
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "rtTriangle");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x7: //parallelogram
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "parallelogram");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x8: //trapezoid
                        //Utils.GetDefaultDocument("shapes.trapezoid").WriteTo(_writer);                         
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "nonIsoscelesTrapezoid");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x9: //hexagon
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "hexagon");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xA: //octagon
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "octagon");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB: //Plus
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "mathPlus");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC: //Star
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "star5");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xD: //Arrow
                    case 0xE: //ThickArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "rightArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xF: //HomePlate
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "homePlate");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x10: //Cube
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "cube");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x11: //Balloon
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "wedgeEllipseCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x12: //Seal
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "star16");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x13: //Arc
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedConnector2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x14: //Line
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "line");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x15: //Plaque
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "plaque");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x16: //Cylinder
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "can");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x17: //Donut
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "donut");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x18: //TextSimple
                    case 0x19: //TextOctagon
                    case 0x1A: //TextHexagon
                    case 0x1B: //TextCurve
                    case 0x1C: //TextWave
                    case 0x1D: //TextRing
                    case 0x1E: //TextOnCurve
                    case 0x1F: //TextOnRing
                        break;
                    case 0x20: //StraightConnector1
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "straightConnector1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x21: //BentConnector2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bentConnector2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x22: //BentConnector3
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bentConnector3");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x23: //BentConnector4
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bentConnector4");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x24: //BentConnector5
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bentConnector5");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x25: //CurvedConnector2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedConnector2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x26: //CurvedConnector3
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedConnector3");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x27: //CurvedConnector4
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedConnector4");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x28: //CurvedConnector5
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedConnector5");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x29: //Callout1
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "callout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x2A: //Callout2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "callout2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x2B: //Callout3
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "callout3");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x2C: //AccentCallout1
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentCallout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x2D: //AccentCallout2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentCallout2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x2E: //AccentCallout3
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentCallout3");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x2F: //BorderCallout1
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "borderCallout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x30: //BorderCallout2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "borderCallout2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x31: //BorderCallout3
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "borderCallout3");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x32: //AccentBorderCallout1
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentborderCallout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x33: //accentBorderCallout2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentborderCallout2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x34: //accentBorderCallout3
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentborderCallout3");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x35: //Ribbon
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "ribbon");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x36: //Ribbon2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "ribbon2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x37: //Chevron
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "chevron");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x38: //Pentagon
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "pentagon");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x39: //noSmoking
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "noSmoking");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x3A: //Seal8
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "star8");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x3B: //Seal16
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "star16");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x3C: //Seal32
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "star32");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x3D: //WedgeRectCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "wedgeRectCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x3E: //WedgeRRectCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "wedgeRoundRectCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x3F: //WedgeEllipseCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "wedgeEllipseCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x40: //Wave
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "wave");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x41: //FolderCorner
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "foldedCorner");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x42: //LeftArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x43: //DownArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "downArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x44: //UpArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "upArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x45: //LeftRightArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftRightArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x46: //UpDownArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "upDownArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x47: //IrregularSeal1
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "irregularSeal1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x48: //IrregularSeal2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "irregularSeal2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x49: //LightningBolt
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "lightningBolt");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x4A: //Heart
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "heart");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x4B: //PictureFrame
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "frame");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x4C: //QuadArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "quadArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x4D: //LeftArrowCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftArrowCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x4E: //RightArrowCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "rightArrowCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x4F: //UpArrowCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "upArrowCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x50: //DownArrowCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "downArrowCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x51: //LeftRightArrowCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftRightArrowCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x52: //UpDownArrowCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "upDownArrowCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x53: //QuadArrowCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "quadArrowCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x54: //Bevel
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bevel");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x55: //LeftBracket
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftBracket");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x56: //RightBracket
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "rightBracket");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x57: //LeftBrace
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftBrace");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x58: //RightBrace
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "rightBrace");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x59: //LeftUpArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftUpArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x5A: //BentUpArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bentUpArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x5B: //BentArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bentArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x5C: //Seal24
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "star24");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x5D: //stripedRightArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "stripedRightArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x5E: //notchedRightArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "notchedRightArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x5F: //BlockArc
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "blockArc");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x60: //SmileyFace
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "smileyFace");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x61: //verticalScroll
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "verticalScroll");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x62: //horizontalScroll
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "horizontalScroll");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x63: //circularArrow
                    case 0x64: //notchedCircularArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "circularArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x65: //uturnArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "uturnArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x66: //curvedRightArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedRightArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x67: //curvedLeftArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedLeftArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x68: //curvedUpArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedUpArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x69: //curvedDownArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "curvedDownArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x6A: //CloudCallout
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "cloudCallout");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x6B: //EllipseRibbon
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "ellipseRibbon");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x6C: //EllipseRibbon2
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "ellipseRibbon2");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x6D: //flowChartProcess
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartProcess");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x6E: //flowChartDecision
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartDecision");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x6F: //flowChartInputOutput
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartInputOutput");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x70: //flowChartPredefinedProcess
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartPredefinedProcess");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x71: //flowChartInternalStorage
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartInternalStorage");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x72: //flowChartDocument
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartDocument");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x73: //flowChartMultidocument
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartMultidocument");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x74: //flowChartTerminator
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartTerminator");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x75: //flowChartPreparation
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartPreparation");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x76: //flowChartManualInput
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartManualInput");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x77: //flowChartManualOperation
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartManualOperation");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x78: //flowChartConnector
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartConnector");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x79: //flowChartPunchedCard
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartPunchedCard");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x7A: //flowChartPunchedTape
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartPunchedTape");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x7B: //flowChartSummingJunction
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartSummingJunction");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x7C: //flowChartOr
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartOr");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x7D: //flowChartCollate
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartCollate");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x7E: //flowChartSort
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartSort");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x7F: //flowChartExtract
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartExtract");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x80: //flowChartMerge
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartMerge");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x81: //flowChartOfflineStorage
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartOfflineStorage");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x82: //flowChartOnlineStorage
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartOnlineStorage");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x83: //flowChartMagneticTape
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartMagneticTape");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x84: //flowChartMagneticDisk
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartMagneticDisk");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x85: //flowChartMagneticDrum
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartMagneticDrum");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x86: //flowChartDisplay
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartDisplay");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x87: //flowChartDelay
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowChartDelay");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0x88: //TextPlaynText
                    case 0x89: //TextStop
                    case 0x8A: //TextTriangle
                    case 0x8B: //TextTriangleInverted
                    case 0x8C: //TextChevron
                    case 0x8D: //TextChevronInverted
                    case 0x8E: //TextRingInside
                    case 0x8F: //TextRingOutside
                    case 0x90: //TextArchUpCurve
                    case 0x91: //TextArchDownCurve
                    case 0x92: //TextCircleCurve
                    case 0x93: //TextButtonCurve
                    case 0x94: //TextArchUpPour
                    case 0x95: //TextArchDownPour
                    case 0x96: //TextCirclePout
                    case 0x97: //TextButtonPout
                    case 0x98: //TextCurveUp
                    case 0x99: //TextCurveDown
                    case 0x9A: //TextCascadeUp
                    case 0x9B: //TextCascadeDown
                    case 0x9C: //TextWave1
                    case 0x9D: //TextWave2
                    case 0x9E: //TextWave3
                    case 0x9F: //TextWave4
                    case 0xA0: //TextInflate
                    case 0xA1: //TextDeflate
                    case 0xA2: //TextInflateBottom
                    case 0xA3: //TextDeflateBottom
                    case 0xA4: //TextInflateTop
                    case 0xA5: //TextDeflateTop
                    case 0xA6: //TextDeflateInflate
                    case 0xA7: //TextDeflateInflateDeflate
                    case 0xA8: //TextFadeRight
                    case 0xA9: //TextFadeLeft
                    case 0xAA: //TextFadeUp
                    case 0xAB: //TextFadeDown
                    case 0xAC: //TextSlantUp
                    case 0xAD: //TextSlantDown
                    case 0xAE: //TextCanUp
                    case 0xAF: //TextCanDown
                        break;
                    case 0xB0: //flowchartAlternateProcess
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowchartAlternateProcess");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB1: //flowchartOffpageConnector
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "flowchartOffpageConnector");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB2: //Callout90
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "callout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB3: //AccentCallout90
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentCallout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB4: //BorderCallout90
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "borderCallout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB5: //AccentBorderCallout90
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "accentBorderCallout1");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB6: //LeftRightUpArrow
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "leftRightUpArrow");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB7: //Sun
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "sun");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB8: //Moon
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "moon");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xB9: //BracketPair
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bracketPair");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xBA: //BracePair
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "bracePair");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xBB: //Seal4
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "star4");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xBC: //DoubleWave
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "doubleWave");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xBD: //ActionButtonBlank
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonBlank");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xBE: //ActionButtonHome
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonHome");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xBF: //ActionButtonHelp
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonHelp");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC0: //ActionButtonInformation
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonInformation");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC1: //ActionButtonForwardNext
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonForwardNext");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC2: //ActionButtonBackPrevious
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonBackPrevious");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC3: //ActionButtonEnd
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonEnd");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC4: //ActionButtonBeginning
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonBeginning");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC5: //ActionButtonReturn
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonReturn");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC6: //ActionButtonDocument
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonDocument");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC7: //ActionButtonSound
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonSound");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC8: //ActionButtonMovie
                        _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "actionButtonMovie");
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement(); //prstGeom
                        break;
                    case 0xC9: //HostControl (do not use)
                        break;
                    case 0xCA: //TextBox
                        break;
                    default:
                        break;
                }
            }
        }

        private Dictionary<int, int> spidToId = new Dictionary<int, int>();
        private void WriteCNvPr(int spid, string name)
        {
            _writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++_idCnt).ToString());
            _writer.WriteAttributeString("name", name);
            _writer.WriteEndElement();
            if (!spidToId.ContainsKey(spid)) spidToId.Add(spid, _idCnt);
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
