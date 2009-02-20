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
    class CharacterRunPropsMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;

        public CharacterRunPropsMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public void Apply(CharacterRun run, string startElement, Slide slide, ref string lastColor)
        {
            //_writer.WriteStartElement("a", "rPr", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", startElement, OpenXmlNamespaces.DrawingML);

            if (run.SizePresent)
            {
                _writer.WriteAttributeString("sz", (run.Size * 100).ToString());
            }

            if (run.StyleFlagsFieldPresent)
            {
                if ((run.Style & StyleMask.IsBold) == StyleMask.IsBold) _writer.WriteAttributeString("b", "1");
                if ((run.Style & StyleMask.IsItalic) == StyleMask.IsItalic) _writer.WriteAttributeString("i", "1");
                if ((run.Style & StyleMask.IsUnderlined) == StyleMask.IsUnderlined) _writer.WriteAttributeString("u", "sng");
            }
            if (run.ColorPresent)
            {
                _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);

                if (run.Color.IsSchemeColor) //TODO: to be fully implemented
                {
                    //_writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);


                    List<ColorSchemeAtom> colors = slide.AllChildrenWithType<ColorSchemeAtom>();
                    ColorSchemeAtom MasterScheme = null;
                    foreach (ColorSchemeAtom color in colors)
                    {
                        if (color.Instance == 1) MasterScheme = color;
                    }

                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    switch (run.Color.Index)
                    {
                        case 0x00: //background
                            lastColor = new RGBColor(MasterScheme.Background, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0x01: //text
                            lastColor =  new RGBColor(MasterScheme.TextAndLines, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0x02: //shadow
                            lastColor =  new RGBColor(MasterScheme.Shadows, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0x03: //title
                            lastColor =  new RGBColor(MasterScheme.TitleText, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0x04: //fill
                            lastColor =  new RGBColor(MasterScheme.Fills, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0x05: //accent1
                            lastColor =  new RGBColor(MasterScheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0x06: //accent2
                            lastColor =  new RGBColor(MasterScheme.AccentAndHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0x07: //accent3
                            lastColor =  new RGBColor(MasterScheme.AccentAndFollowedHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0xFE: //sRGB
                            lastColor =  run.Color.Red.ToString("X").PadLeft(2, '0') + run.Color.Green.ToString("X").PadLeft(2, '0') + run.Color.Blue.ToString("X").PadLeft(2, '0');
                            _writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0xFF: //undefined
                            break;
                    }
                    _writer.WriteEndElement();
                    //_writer.WriteEndElement();
                }
                else
                {
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    lastColor = run.Color.Red.ToString("X").PadLeft(2, '0') + run.Color.Green.ToString("X").PadLeft(2, '0') + run.Color.Blue.ToString("X").PadLeft(2, '0');
                    _writer.WriteAttributeString("val", lastColor);
                    _writer.WriteEndElement();
                }
                _writer.WriteEndElement();
            }
            else if (lastColor.Length > 0)
            {
                _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("val",lastColor);
                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }


            if (run.StyleFlagsFieldPresent)
            {
                if ((run.Style & StyleMask.HasShadow) == StyleMask.HasShadow)
                {
                    //TODO: these values are default and have to be replaced
                    _writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("blurRad", "38100");
                    _writer.WriteAttributeString("dist", "38100");
                    _writer.WriteAttributeString("dir", "2700000");
                    _writer.WriteAttributeString("algn", "tl");
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", "C0C0C0");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                }

                if ((run.Style & StyleMask.IsEmbossed) == StyleMask.IsEmbossed)
                {
                    //TODO: these values are default and have to be replaced
                    _writer.WriteStartElement("a", "effectDag", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("name", "");
                    _writer.WriteStartElement("a", "cont", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("type", "tree");
                    _writer.WriteAttributeString("name", "");
                    _writer.WriteStartElement("a", "effect", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("ref", "fillLine");
                    _writer.WriteEndElement();
                    _writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("dist", "38100");
                    _writer.WriteAttributeString("dir", "13500000");
                    _writer.WriteAttributeString("algn", "br");
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", "FFFFFF");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    _writer.WriteStartElement("a", "cont", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("type", "tree");
                    _writer.WriteAttributeString("name", "");
                    _writer.WriteStartElement("a", "effect", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("ref", "fillLine");
                    _writer.WriteEndElement();
                    _writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("dist", "38100");
                    _writer.WriteAttributeString("dir", "2700000");
                    _writer.WriteAttributeString("algn", "tl");
                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", "999999");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    _writer.WriteStartElement("a", "effect", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("ref", "fillLine");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                }

                //TODOS
                //HasAsianSmartQuotes 
                //HasHorizonNumRendering 
                //ExtensionNibble 

            }

            //TODOs:
            //run.ANSITypefacePresent
            //run.FEOldTypefacePresent
            //run.PositionPresent
            //run.SymbolTypefacePresent
            //run.TypefacePresent

            if (run.TypefacePresent)
            {
                //List<ColorSchemeAtom> colors = slide.AllChildrenWithType<ColorSchemeAtom>();
                //ColorSchemeAtom MasterScheme = null;
                //foreach (ColorSchemeAtom color in colors)
                //{
                //    if (color.Instance == 1) MasterScheme = color;
                //}


                _writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
                try
                {
                    FontCollection fonts = _ctx.Ppt.DocumentRecord.FirstChildWithType<DIaLOGIKa.b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                    FontEntityAtom entity = fonts.entities[(int)run.TypefaceIdx];
                    if (entity.TypeFace.IndexOf('\0') > 0)
                    {
                        _writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                    }
                    else
                    {
                        _writer.WriteAttributeString("typeface", entity.TypeFace);
                    }
                }
                catch (Exception ex)
                {
                    throw;
                }

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
        }
    }
}
