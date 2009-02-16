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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Reflection;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class MasterMapping : PresentationMapping<Slide>
    {
        public SlideMasterPart MasterPart;
        protected Slide Master;
        protected UInt32 MasterId;
        protected MasterLayoutManager LayoutManager;

        public MasterMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddSlideMasterPart())
        {
            this.MasterPart = (SlideMasterPart)this.targetPart;
        }

        override public void Apply(Slide master)
        {
            TraceLogger.DebugInternal("MasterMapping.Apply");
            UInt32 masterId = master.PersistAtom.SlideId;
            _ctx.RegisterMasterMapping(masterId, this);

            this.Master = master;
            this.MasterId = master.PersistAtom.SlideId;
            this.LayoutManager = _ctx.GetOrCreateLayoutManagerByMasterId(this.MasterId);

            // Add PPT2007 roundtrip slide layouts
            List<RoundTripContentMasterInfo12> rtSlideLayouts = this.Master.AllChildrenWithType<RoundTripContentMasterInfo12>();

            foreach (RoundTripContentMasterInfo12 slideLayout in rtSlideLayouts)
            {
                SlideLayoutPart layoutPart = this.LayoutManager.AddLayoutPartWithInstanceId(slideLayout.Instance);

                slideLayout.XmlDocumentElement.WriteTo(layoutPart.XmlWriter);
                layoutPart.XmlWriter.Flush();
            }
        }

        public void Write()
        {
            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "sldMaster", OpenXmlNamespaces.PresentationML);
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            _writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);

            //TODO: write p:bg
            ShapeContainer sc = this.Master.FirstChildWithType<PPDrawing>().FirstChildWithType<DrawingContainer>().FirstChildWithType<ShapeContainer>();
            if (sc != null)
            {
                Shape sh = sc.FirstChildWithType<Shape>();
                ShapeOptions so = sc.FirstChildWithType<ShapeOptions>();
                int indexOfPicture = -1;
                string name;
                if (so != null)
                {
                    foreach (ShapeOptions.OptionEntry en in so.Options)
                    {
                        switch (en.pid)
                        {
                            case ShapeOptions.PropertyId.fillBlip:
                                indexOfPicture = ((int)en.op) - 1;
                                break;
                            case ShapeOptions.PropertyId.fillBlipName:
                                name = Encoding.Unicode.GetString(en.opComplex);
                                name = name.Substring(0, name.Length - 1);
                                break;
                            default:
                                break;

                        }
                    }
                }
                if (indexOfPicture > -1)
                {
                    DrawingGroup gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                    BlipStoreEntry bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)indexOfPicture];
                    string rId;

                    //if (this.parentSlideMapping is MasterMapping) return;

                    if (this._ctx.AddedImages.ContainsKey(bse.foDelay))
                    {
                        rId = this._ctx.AddedImages[bse.foDelay];
                    }
                    else
                    {

                        if (!_ctx.Ppt.PicturesContainer._pictures.ContainsKey(bse.foDelay))
                        {
                            return;
                        }

                        BitmapBlip b = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                        ImagePart imgPart = null;
                        imgPart = this.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                        imgPart.TargetDirectory = "..\\media";
                        System.IO.Stream outStream = imgPart.GetStream();
                        outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);

                        rId = imgPart.RelIdToString;
                        //this._ctx.AddedImages.Add(bse.foDelay, rId);
                    }

                    _writer.WriteStartElement("p", "bg", OpenXmlNamespaces.PresentationML);

                    _writer.WriteStartElement("p", "bgPr", OpenXmlNamespaces.PresentationML);

                    _writer.WriteStartElement("a", "blipFill", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("embed", OpenXmlNamespaces.Relationships, rId);
                    _writer.WriteEndElement(); //a:blip
                    _writer.WriteElementString("a", "srcRect", OpenXmlNamespaces.DrawingML, "");
                    _writer.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingML);
                    _writer.WriteElementString("a", "fillRect", OpenXmlNamespaces.DrawingML, "");
                    _writer.WriteEndElement(); //a:stretch
                    _writer.WriteEndElement(); //p:blipFill

                    _writer.WriteElementString("a", "effectLst", OpenXmlNamespaces.DrawingML, "");

                    _writer.WriteEndElement(); //p:bgPr

                    _writer.WriteEndElement(); //p:bg
                }
                else
                {
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                    {
                        switch (so.OptionsByID[ShapeOptions.PropertyId.fillType].op)
                        {
                            case 0:
                                //no background
                                break;
                            default:
                                _writer.WriteStartElement("p", "bg", OpenXmlNamespaces.PresentationML);
                                _writer.WriteStartElement("p", "bgPr", OpenXmlNamespaces.PresentationML);
                                new FillMapping(_ctx, _writer, this).Apply(so);
                                _writer.WriteEndElement();
                                _writer.WriteEndElement();
                                break;
                        }
                    }
                }
            }

            _writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);
            ShapeTreeMapping stm = new ShapeTreeMapping(_ctx, _writer);
            stm.parentSlideMapping = this;
            stm.Apply(this.Master.FirstChildWithType<PPDrawing>());

            _writer.WriteEndElement();
            _writer.WriteEndElement();

            // Write clrMap
            ColorMappingAtom clrMap = this.Master.FirstChildWithType<ColorMappingAtom>();
            if (clrMap != null)
            {
                // clrMap from ColorMappingAtom wrongly uses namespace DrawingML
                _writer.WriteStartElement("p", "clrMap", OpenXmlNamespaces.PresentationML);

                foreach (XmlAttribute attr in clrMap.XmlDocumentElement.Attributes)
                    if (attr.Prefix != "xmlns")
                        _writer.WriteAttributeString(attr.LocalName, attr.Value);

                _writer.WriteEndElement();
            }
            else
            {
                // In absence of ColorMappingAtom write default clrMap
                Utils.GetDefaultDocument("clrMap").WriteTo(_writer);
            }

            // Write slide layout part id list
            _writer.WriteStartElement("p", "sldLayoutIdLst", OpenXmlNamespaces.PresentationML);

            List<SlideLayoutPart> layoutParts = this.LayoutManager.GetAllLayoutParts();

            // Master must have at least one SlideLayout or RepairDialog will appear
            if (layoutParts.Count == 0)
            {
                SlideLayoutPart layoutPart = this.LayoutManager.GetOrCreateLayoutPartByLayoutType(0, null);
                layoutParts.Add(layoutPart);
            }

            foreach (SlideLayoutPart slideLayoutPart in layoutParts)
            {
                _writer.WriteStartElement("p", "sldLayoutId", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, slideLayoutPart.RelIdToString);
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();

            // Write txStyles
            RoundTripOArtTextStyles12 roundTripTxStyles = this.Master.FirstChildWithType<RoundTripOArtTextStyles12>();
            if (roundTripTxStyles != null)
            {
                roundTripTxStyles.XmlDocumentElement.WriteTo(_writer);
            }
            else
            {
                //throw new NotImplementedException("Write txStyles in case of PPT without roundTripTxStyles"); // TODO (pre PP2007)
                
                //XmlDocument slideLayoutDoc = Utils.GetDefaultDocument("txStyles");
                //slideLayoutDoc.WriteTo(_writer);

                new TextMasterStyleMapping(_ctx, _writer, this).Apply(this.Master);
            }

            // Write theme
            //
            // Note: We need to create a new theme part for each master,
            // even if it they have the same content.
            //
            // Otherwise PPT will complain about the structure of the file.
            ThemePart themePart = _ctx.Pptx.PresentationPart.AddThemePart();

            XmlNode xmlDoc;
            Theme theme = this.Master.FirstChildWithType<Theme>();

            if (theme != null)
            {
                xmlDoc = theme.XmlDocumentElement;
                xmlDoc.WriteTo(themePart.XmlWriter);
            }
            else
            {
                List<ColorSchemeAtom> schemes = this.Master.AllChildrenWithType<ColorSchemeAtom>();
                if (schemes.Count > 0)
                {
                    new ColorSchemeMapping(_ctx, themePart.XmlWriter).Apply(schemes);                    
                }
                else
                {
                    // In absence of Theme record use default theme
                    xmlDoc = Utils.GetDefaultDocument("theme");
                    xmlDoc.WriteTo(themePart.XmlWriter);
                }
            }

            themePart.XmlWriter.Flush();
           

            this.MasterPart.ReferencePart(themePart);
            
            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
