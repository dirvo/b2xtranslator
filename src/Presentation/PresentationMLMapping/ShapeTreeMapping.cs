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
    class ShapeTreeMapping :
        AbstractOpenXmlMapping,
        IMapping<PPDrawing>
    {
        protected int _idCnt;
        protected ConversionContext _ctx;
        protected string _footertext;

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

        //used to give each group a unique identifyer
        private int groupcounter = -10;
        public void Apply(GroupContainer group)
        {
            _writer.WriteStartElement("p", "grpSp", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "nvGrpSpPr", OpenXmlNamespaces.PresentationML);
            WriteCNvPr(--groupcounter, "");
            _writer.WriteElementString("p", "cNvGrpSpPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteEndElement(); //nvGrpSpPr

            _writer.WriteStartElement("p", "grpSpPr", OpenXmlNamespaces.PresentationML);
            GroupShapeRecord gsr = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>();
            ClientAnchor anchor = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();

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
                _writer.WriteAttributeString("x", gsr.rcgBounds.Left.ToString());
                _writer.WriteAttributeString("y", gsr.rcgBounds.Top.ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "chExt", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("cx", (gsr.rcgBounds.Right - gsr.rcgBounds.Left).ToString());
                _writer.WriteAttributeString("cy", (gsr.rcgBounds.Bottom - gsr.rcgBounds.Top).ToString());
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


        private ShapeOptions so;
        public void Apply(ShapeContainer container)
        {
            Apply(container, "");
        }
        public void Apply(ShapeContainer container, string footertext)
        {
            _footertext = footertext;
            ClientData clientData = container.FirstChildWithType<ClientData>();

            bool continueShape = true;

            Shape sh = container.FirstChildWithType<Shape>();
            so = container.FirstChildWithType<ShapeOptions>();
            //if (clientData == null)
                if (so != null)
                {
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.Pib))
                    {
                        writePic(container);
                        continueShape = false;
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
                if (false & sndSo.OptionsByID.ContainsKey(ShapeOptions.PropertyId.metroBlob))
                {
                    ZipUtils.ZipReader reader = null;
                    try
                    {
                        ShapeOptions.OptionEntry metroBlob = sndSo.OptionsByID[ShapeOptions.PropertyId.metroBlob];
                        byte[] code = metroBlob.opComplex;
                        string path = System.IO.Path.GetTempFileName();
                        System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Create);
                        fs.Write(code, 0, code.Length);
                        fs.Close();

                        reader = ZipUtils.ZipFactory.OpenArchive(path);
                        System.IO.StreamReader mems = new System.IO.StreamReader(reader.GetEntry("drs/shapexml.xml"));
                        string xml = mems.ReadToEnd();
                        xml = xml.Substring(xml.IndexOf("<p:sp")); //remove xml declaration

                        _writer.WriteRaw(xml);

                        continueShape = false;

                        reader.Close();
                    }
                    catch (Exception e)
                    {
                        continueShape = true;
                        if (reader != null) reader.Close();
                    }
                }
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
                            case 3011:
                                placeholder = (OEPlaceHolderAtom)rec; // clientData.FirstChildWithType<OEPlaceHolderAtom>();

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
                            default:
                                break;
                        }
                    }
                }

                _writer.WriteEndElement();

                _writer.WriteEndElement();


                // Visible shape properties
                _writer.WriteStartElement("p", "spPr", OpenXmlNamespaces.PresentationML);

                ClientAnchor anchor = container.FirstChildWithType<ClientAnchor>();
                ChildAnchor chAnchor = container.FirstChildWithType<ChildAnchor>();

                if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
                {
                    _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);
                    if (sh.fFlipH) _writer.WriteAttributeString("flipH", "1");
                    if (sh.fFlipV) _writer.WriteAttributeString("flipV", "1");

                    if (container.FirstAncestorWithType<GroupContainer>().FirstAncestorWithType<GroupContainer>() == null)
                    {
                        _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                        _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                        _writer.WriteEndElement();

                        _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                        _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                        _writer.WriteEndElement();
                    }
                    else
                    {
                        _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("x", (anchor.Left).ToString());
                        _writer.WriteAttributeString("y", (anchor.Top).ToString());
                        _writer.WriteEndElement();

                        _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("cx", (anchor.Right - anchor.Left).ToString());
                        _writer.WriteAttributeString("cy", (anchor.Bottom - anchor.Top).ToString());
                        _writer.WriteEndElement();
                    }
                    _writer.WriteEndElement();
                }
                else if (chAnchor != null && chAnchor.Right >= chAnchor.Left && chAnchor.Bottom >= chAnchor.Top)
                {
                    ClientAnchor groupAnchor = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();
                    Rectangle rec = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>().rcgBounds;

                    _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);
                    if (sh.fFlipH) _writer.WriteAttributeString("flipH", "1");
                    if (sh.fFlipV) _writer.WriteAttributeString("flipV", "1");

                    _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("x", chAnchor.Left.ToString());
                    _writer.WriteAttributeString("y", chAnchor.Top.ToString());
                    _writer.WriteEndElement();

                    _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("cx", (chAnchor.Right - chAnchor.Left).ToString());
                    _writer.WriteAttributeString("cy", (chAnchor.Bottom - chAnchor.Top).ToString());
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

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                {
                    new FillMapping(_ctx, _writer, parentSlideMapping).Apply(so);
                }
                else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                {
                    if (sh.Instance != 0xca & placeholder == null)
                    {
                        string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, container.FirstAncestorWithType<Slide>(), so);
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
                            } else {
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
                            string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, container.FirstAncestorWithType<Slide>(), so);
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
                    writeBodyPr();
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
                    //    name = Encoding.Unicode.GetString(en.opComplex);
                    //    name = name.Substring(0, name.Length - 1).Replace("\0","");
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

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureBrightness) | so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureContrast))
            {
                _writer.WriteStartElement("a", "lum", OpenXmlNamespaces.DrawingML);

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureBrightness))
                {
                    UInt32 b = so.OptionsByID[ShapeOptions.PropertyId.pictureBrightness].op;
                    if (b == 0xFFF8000) b = 0;
                    Decimal b1 = (Decimal)b / 0x8000;
                    b1 = b1 * 100000;
                    b1 = Math.Floor(b1);
                    _writer.WriteAttributeString("bright", b1.ToString());
                }

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureContrast))
                {
                    UInt32 b = so.OptionsByID[ShapeOptions.PropertyId.pictureContrast].op;
                    if (b == 0x7FFFFFFF) b = 0;
                    Decimal b2 = (Decimal)b / 0x10000;
                    b2 = b2 - 1;
                    b2 = b2 * 100000;
                    b2 = Math.Floor(b2);
                    _writer.WriteAttributeString("contrast", b2.ToString());
                }

                _writer.WriteEndElement();
            }
           

            _writer.WriteEndElement(); //a:blip
            _writer.WriteStartElement("a", "srcRect", OpenXmlNamespaces.DrawingML);
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromLeft))
            {
                _writer.WriteAttributeString("l", Math.Floor((Decimal)so.OptionsByID[ShapeOptions.PropertyId.cropFromLeft].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromTop))
            {
                _writer.WriteAttributeString("t",Math.Floor((Decimal)so.OptionsByID[ShapeOptions.PropertyId.cropFromTop].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromRight))
            {
                _writer.WriteAttributeString("r", Math.Floor((Decimal)so.OptionsByID[ShapeOptions.PropertyId.cropFromRight].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromBottom))
            {
                _writer.WriteAttributeString("b", Math.Floor((Decimal)so.OptionsByID[ShapeOptions.PropertyId.cropFromBottom].op / 65536 * 100000).ToString());
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

            _writer.WriteEndElement(); //p:spPr

            

            _writer.WriteEndElement(); //p:pic
        }

        public void writeBodyPr()
        {
            _writer.WriteStartElement("a", "bodyPr", OpenXmlNamespaces.DrawingML);

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

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fFitTextToShape))
            {
                _writer.WriteElementString("a", "spAutoFit", OpenXmlNamespaces.DrawingML, "");
            }

            _writer.WriteEndElement();
        }

        public void Apply(ClientTextbox textbox)
        {
            _writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);

            writeBodyPr();

            _writer.WriteStartElement("a", "lstStyle", OpenXmlNamespaces.DrawingML);

            System.IO.MemoryStream ms = new System.IO.MemoryStream(textbox.Bytes);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            List<int> lst = new List<int>();
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

            new TextMapping(_ctx, _writer).Apply(textbox, _footertext);

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
                    } else {
                        _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                    }
                    _writer.WriteEndElement(); //prstGeom
                }                
            }
        }

        

        private Dictionary<int, int> spidToId = new Dictionary<int, int>();
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
