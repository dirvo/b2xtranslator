﻿/*
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
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class FillMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;
        protected PresentationMapping<RegularContainer> _parentSlideMapping;

        public FillMapping(ConversionContext ctx, XmlWriter writer, PresentationMapping<RegularContainer> parentSlideMapping)
            : base(writer)
        {
            _ctx = ctx;
            _parentSlideMapping = parentSlideMapping;
        }

        public void Apply(ShapeOptions so)
        {
            RegularContainer slide = so.FirstAncestorWithType<Slide>();
            if (slide == null) slide = so.FirstAncestorWithType<Note>();
            string colorval = "";
            switch (so.OptionsByID[ShapeOptions.PropertyId.fillType].op)
            {
                case 0x0: //solid
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, (RegularContainer)slide, so);
                    } else {
                        colorval = "000000"; //TODO: find out which color to use in this case
                    }
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
                    break;
                case 0x1: //pattern
                    uint blipIndex1 = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    DrawingGroup gr1 = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                    BlipStoreEntry bse1 = (BlipStoreEntry)gr1.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex1 - 1];
                    BitmapBlip b1 = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse1.foDelay];

                    _writer.WriteStartElement("a", "pattFill", OpenXmlNamespaces.DrawingML);

                    _writer.WriteAttributeString("prst", Utils.getPrstForPatternCode(b1.m_bTag)); //Utils.getPrstForPattern(blipNamePattern));

                    _writer.WriteStartElement("a", "fgClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide,so));
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();

                    _writer.WriteStartElement("a", "bgClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    }
                    else
                    {
                        colorval = "ffffff"; //TODO: find out which color to use in this case
                    }
                    _writer.WriteAttributeString("val", colorval);
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();

                    _writer.WriteEndElement();

                    break;
                case 0x2: //texture
                case 0x3: //picture
                    uint blipIndex = 0;
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBlip))
                    {
                        blipIndex = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    } else 
                    {  
                        blipIndex = so.OptionsByID[ShapeOptions.PropertyId.Pib].op;
                    }

                    //string blipName = Encoding.UTF8.GetString(so.OptionsByID[ShapeOptions.PropertyId.fillBlipName].opComplex);
                    string rId = "";
                    DrawingGroup gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];

                    if (blipIndex < gr.FirstChildWithType<BlipStoreContainer>().Children.Count)
                    {
                        BlipStoreEntry bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex - 1];

                        if (_ctx.Ppt.PicturesContainer._pictures.ContainsKey(bse.foDelay))
                        {

                            BitmapBlip b = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                            ImagePart imgPart = null;
                            imgPart = _parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                            imgPart.TargetDirectory = "..\\media";
                            System.IO.Stream outStream = imgPart.GetStream();
                            outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);

                            rId = imgPart.RelIdToString;

                            _writer.WriteStartElement("a", "blipFill", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("dpi", "0");
                            _writer.WriteAttributeString("rotWithShape", "1");

                            _writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, rId);

                           

                            _writer.WriteEndElement();

                            _writer.WriteElementString("a", "srcRect", OpenXmlNamespaces.DrawingML, "");

                            if (so.OptionsByID[ShapeOptions.PropertyId.fillType].op == 0x3)
                            {
                                _writer.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingML);
                                _writer.WriteElementString("a", "fillRect", OpenXmlNamespaces.DrawingML, "");
                                _writer.WriteEndElement();
                            }
                            else
                            {
                                _writer.WriteStartElement("a", "tile", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("tx", "0");
                                _writer.WriteAttributeString("ty", "0");
                                _writer.WriteAttributeString("sx", "100000");
                                _writer.WriteAttributeString("sy", "100000");
                                _writer.WriteAttributeString("flip", "none");
                                _writer.WriteAttributeString("algn", "tl");
                                _writer.WriteEndElement();
                            }

                            _writer.WriteEndElement();
                        }
                    }
                    break;
                case 0x4: //shade
                case 0x5: //shadecenter
                case 0x6: //shadeshape
                    _writer.WriteStartElement("a", "gradFill", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("rotWithShape", "1");
                    _writer.WriteStartElement("a", "gsLst", OpenXmlNamespaces.DrawingML);
                    bool useFillAndBack = true;

                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillShadeColors))
                    {

                        byte[] colors = so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].opComplex;

                        if (colors.Length > 0)
                        {

                            useFillAndBack = false;
                            ShapeOptions.OptionEntry type = so.OptionsByID[ShapeOptions.PropertyId.fillShadeType];

                            UInt16 nElems = System.BitConverter.ToUInt16(colors, 0);
                            UInt16 nElemsAlloc = System.BitConverter.ToUInt16(colors, 2);
                            UInt16 cbElem = System.BitConverter.ToUInt16(colors, 4);

                            List<string> positions = new List<string>();

                            switch (nElems)
                            {
                                case 1:
                                case 2:
                                case 3:
                                case 4:
                                case 5:
                                    positions.Add("0");
                                    positions.Add("30000");
                                    positions.Add("65000");
                                    positions.Add("90000");
                                    positions.Add("100000");
                                    break;
                                case 6:
                                case 7:
                                case 8:
                                case 9:
                                case 10:
                                default:
                                    positions.Add("0");
                                    positions.Add("8000");
                                    positions.Add("13000");
                                    positions.Add("21000");
                                    positions.Add("52000");
                                    positions.Add("56000");
                                    positions.Add("58000");
                                    positions.Add("71000");
                                    positions.Add("94000");
                                    positions.Add("100000");
                                    break;
                            }


                            string[] alphas = new string[nElems];
                            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                            {
                                decimal end = Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000));
                                decimal start = Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000));
                                alphas[0] = start.ToString();
                                for (int i = 1; i < nElems - 1; i++)
                                {
                                    alphas[i] = Math.Round(start + (end - start) / 3 * i).ToString();
                                }
                                //alphas[1] = Math.Round(start + (end - start) / 3).ToString();
                                //alphas[2] = Math.Round(start + (end - start) / 3 * 2).ToString();
                                //alphas[3] = Math.Round(start + (end - start) / 3 * 3).ToString();
                                alphas[nElems - 1] = end.ToString();
                            }

                            for (int i = 0; i < nElems * cbElem; i += cbElem)
                            {
                                colorval = Utils.getRGBColorFromOfficeArtCOLORREF(System.BitConverter.ToUInt32(colors, 6 + i), slide, so);
                                _writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("pos", positions[i / cbElem]);

                                _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("val", colorval);
                                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                                {
                                    _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", alphas[i / cbElem]); //we need the percentage of the opacity (65536 means 100%)
                                    _writer.WriteEndElement();
                                }
                                _writer.WriteEndElement();

                                _writer.WriteEndElement();
                            }
                        }
                    }
                    
                    if (useFillAndBack)
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);

                        _writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("pos", "0");
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

                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);

                        _writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("pos", "100000");
                        _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", colorval);
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackOpacity))
                        {
                            _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            _writer.WriteEndElement();
                        }
                        _writer.WriteEndElement();
                        _writer.WriteEndElement();
                    }

                    _writer.WriteEndElement(); //gsLst

                    switch (so.OptionsByID[ShapeOptions.PropertyId.fillType].op)
                    {
                        case 0x6:
                            _writer.WriteStartElement("a", "path", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("path", "shape");
                            _writer.WriteStartElement("a", "fillToRect", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("l", "50000");
                            _writer.WriteAttributeString("t", "50000");
                            _writer.WriteAttributeString("r", "50000");
                            _writer.WriteAttributeString("b", "50000");
                            _writer.WriteEndElement();
                            _writer.WriteEndElement(); //path
                            break;
                        default:
                            _writer.WriteStartElement("a", "path", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("path", "rect");
                            _writer.WriteStartElement("a", "fillToRect", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("r", "100000");
                            _writer.WriteAttributeString("b", "100000");
                            _writer.WriteEndElement();
                            _writer.WriteEndElement(); //path
                            break;
                    }

                    _writer.WriteEndElement(); //gradFill

                    break;
                case 0x7: //shadescale
                    _writer.WriteStartElement("a", "gradFill", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("rotWithShape", "1");
                    _writer.WriteStartElement("a", "gsLst", OpenXmlNamespaces.DrawingML);

                    decimal angle = 90;
                    bool switchColors = false;
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillAngle))
                    {
                        if (so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op != 0)
                        {
                            byte[] bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op);
                            int integral = BitConverter.ToInt16(bytes, 0);
                            uint fractional = BitConverter.ToUInt16(bytes, 2);
                            Decimal result = integral + ((decimal)fractional / (decimal)65536);
                            angle = 65536 - fractional; //I have no idea why this works!!                    
                            angle = angle - 90;
                            if (angle < 0)
                            {
                                angle += 360;
                                switchColors = true;
                            }
                        }
                    }

                    if (switchColors & so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    }
                    else
                    {
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                        {
                            colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);
                        } else {
                            colorval = "000000"; //TODO: find out which color to use in this case
                        }
                    }

                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillShadeType))
                    {
                        //TODO
                    }

                    _writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("pos", "0");
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

                    if (switchColors | !so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);
                    }
                    else
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    }
                    _writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("pos", "100000");
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", colorval);
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackOpacity))
                    {
                        _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                        _writer.WriteEndElement();
                    }
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();

                    _writer.WriteEndElement(); //gsLst

                    _writer.WriteStartElement("a", "lin", OpenXmlNamespaces.DrawingML);

                    _writer.WriteAttributeString("ang", (angle * 60000).ToString());
                    _writer.WriteAttributeString("scaled", "1");
                    _writer.WriteEndElement();

                    _writer.WriteEndElement();
                    break;
                case 0x8: //shadetitle
                case 0x9: //background
                    break;

            }
        }
    }
}
