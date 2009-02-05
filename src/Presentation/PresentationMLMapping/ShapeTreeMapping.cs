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
                foreach (ShapeOptions.OptionEntry en in so.Options)
                {
                    if (en.pid == ShapeOptions.PropertyId.Pib)
                    {
                        if (en.opComplex != null)
                        {
                            //TODO
                        }
                        else
                        {
                            writePic(container);
                            pictureWritten = true;
                        }
                    }
                }
            }

            if (!pictureWritten)
            {

                _writer.WriteStartElement("p", "sp", OpenXmlNamespaces.PresentationML);

                _writer.WriteStartElement("p", "nvSpPr", OpenXmlNamespaces.PresentationML);

                WriteCNvPr("");

                _writer.WriteElementString("p", "cNvSpPr", OpenXmlNamespaces.PresentationML, "");
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

                _writer.WriteEndElement();

                // Descend into unsupported records
                foreach (Record record in container.Children)
                {
                    DynamicApply(record);
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
            //////////////////////////////////////////// TODO...

            //ShapeOptions sos = textbox.FirstAncestorWithType<ShapeContainer>().FirstChildWithType<ShapeOptions>();
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

           

            //////////////////////////

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
            WriteCNvPr("");

            _writer.WriteElementString("p", "cNvGrpSpPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");

            _writer.WriteEndElement();


            // Write visible Group Shape properties
            _writer.WriteStartElement("p", "grpSpPr", OpenXmlNamespaces.PresentationML);
            WriteXFrm(_writer, new Rectangle()); // groupShape.rcgBounds
            _writer.WriteEndElement();
        }

        private void WriteCNvPr(string name)
        {
            _writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++_idCnt).ToString());
            _writer.WriteAttributeString("name", name);
            _writer.WriteEndElement();
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
