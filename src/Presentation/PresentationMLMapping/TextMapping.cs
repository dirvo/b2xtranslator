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
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class TextMapping :
        AbstractOpenXmlMapping,
        IMapping<ClientTextbox>
    {
        protected ConversionContext _ctx;

        public TextMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        /// <summary>
        /// Returns the ParagraphRun of the given style that is active at the given index.
        /// </summary>
        /// <param name="style">style to use</param>
        /// <param name="forIdx">index to use</param>
        /// <returns>ParagraphRun or null in case no run was found</returns>
        protected static ParagraphRun GetParagraphRun(TextStyleAtom style, uint forIdx)
        {
            if (style == null)
                return null;

            uint idx = 0;

            foreach (ParagraphRun p in style.PRuns)
            {
                if (forIdx < idx + p.Length)
                    return p;

                idx += p.Length;
            }

            return null;
        }

        /// <summary>
        /// Returns the CharacterRun of the given style that is active at the given index.
        /// </summary>
        /// <param name="style">style to use</param>
        /// <param name="forIdx">index to use</param>
        /// <returns>CharacterRun or null in case no run was found</returns>
        protected static CharacterRun GetCharacterRun(TextStyleAtom style, uint forIdx)
        {
            if (style == null)
                return null;

            uint idx = 0;

            foreach (CharacterRun c in style.CRuns)
            {
                if (forIdx < idx + c.Length)
                    return c;

                idx += c.Length;
            }

            return null;
        }

        protected static uint GetCharacterRunStart(TextStyleAtom style, uint forIdx)
        {
            if (style == null)
                return 0;

            uint idx = 0;

            foreach (CharacterRun c in style.CRuns)
            {
                if (forIdx < idx + c.Length)
                    return idx;

                idx += c.Length;
            }

            return 0;
        }

        public void Apply(ClientTextbox textbox)
        {
            Apply(textbox, "");
        }

        public void Apply(ClientTextbox textbox, string footertext)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream(textbox.Bytes);
            Record rec = Record.ReadRecord(ms, 0);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            FooterMCAtom mca = null;
            MasterTextPropAtom masterTextProp = null;
            string text = "";

            switch (rec.TypeCode)
            {
                case 3999:
                    thAtom = (TextHeaderAtom)rec;
                    while (ms.Position < ms.Length)
                    {
                        rec = Record.ReadRecord(ms, 0);

                        switch (rec.TypeCode)
                        {
                            case 0xfa0: //TextCharsAtom
                                text = ((TextCharsAtom)rec).Text;
                                thAtom.TextAtom = (TextAtom)rec;
                                break;
                            case 0xfa1: //TextRunStyleAtom
                                style = (TextStyleAtom)rec;
                                style.TextHeaderAtom = thAtom;
                                break;
                            case 0xfa8: //TextBytesAtom
                                text = ((TextBytesAtom)rec).Text;
                                thAtom.TextAtom = (TextAtom)rec;
                                break;
                            case 0xfaa: //TextSpecialInfoAtom
                                //TODO
                                break;
                            case 0xfa2: //MasterTextPropAtom
                                masterTextProp = (MasterTextPropAtom)rec;
                                break;
                            case 0xfd8: //SlideNumberMCAtom
                                SlideNumberMCAtom snmca = (SlideNumberMCAtom)rec;

                                _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

                                _writer.WriteStartElement("a", "fld", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("id", "{18109A10-03E4-4BE3-B6BB-0FCEF851AF87}");
                                _writer.WriteAttributeString("type", "slidenum");
                                _writer.WriteElementString("a", "t", OpenXmlNamespaces.DrawingML, "<#>");
                                _writer.WriteEndElement(); //fld                                
                                _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                                _writer.WriteEndElement(); //endParaRPr
                                _writer.WriteEndElement(); //p

                                text = text.Replace(text.Substring(snmca.Position, 1), "");

                                break;
                            case 0xffa: //FooterMCAtom
                                mca = (FooterMCAtom)rec;
                                text = text.Replace(text.Substring(mca.Position, 1), footertext);

                                foreach (CharacterRun run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            case 0xff8: //GenericDateMCAtom
                                GenericDateMCAtom gdmca = (GenericDateMCAtom)rec;
                                break;
                            default:
                                //TextAtom textAtom = thAtom.TextAtom;
                                //text = (textAtom == null) ? "" : textAtom.Text;
                                //style = thAtom.TextStyleAtom;
                                break;
                        }
                    }
                    break;
                case 3998:
                    OutlineTextRefAtom otrAtom = (OutlineTextRefAtom)rec;
                    SlideListWithText slideListWithText = _ctx.Ppt.DocumentRecord.RegularSlideListWithText;
                                  
                    List<TextHeaderAtom> thAtoms = slideListWithText.SlideToPlaceholderTextHeaders[textbox.FirstAncestorWithType<Slide>().PersistAtom];
                    thAtom = thAtoms[otrAtom.Index];

                    if (thAtom.TextAtom != null) text = thAtom.TextAtom.Text;
                    if (thAtom.TextStyleAtom != null) style = thAtom.TextStyleAtom;

                    break;
                default:
                    throw new NotSupportedException("Can't find text for ClientTextbox without TextHeaderAtom and OutlineTextRefAtom");
            }

            uint idx = 0;

            if (text.Length == 0)
            {
                _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);
                _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                // TODO...
                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }
            else
            {
                String[] parlines = text.Split(new char[] { '\r' }); //text.Split(new char[] { '\v', '\r' });
                int internalOffset = 0;
                foreach (String parline in parlines)
                {
                    String[] runlines = parline.Split(new char[] { '\v' });

                    //each parline forms a paragraph
                    //each runline forms a run

                    ParagraphRun p = GetParagraphRun(style, idx);
                    String runText;
                    writeP(p);

                    uint CharacterRunStart;
                    int len;

                    bool first = true;
                    
                    foreach (string line in runlines)
                    {
                        uint offset = idx;

                        if (!first)
                        {
                            _writer.WriteElementString("a", "br", OpenXmlNamespaces.DrawingML, "");
                            internalOffset += 1;
                        }

                        while (idx < offset + line.Length)
                        {
                            len = line.Length;
                            CharacterRun r = GetCharacterRun(style, idx + (uint)internalOffset + 1);
                            if (r != null)
                            {
                                CharacterRunStart = GetCharacterRunStart(style, idx + (uint)internalOffset + 1);
                                len = (int)(CharacterRunStart + r.Length - idx - internalOffset);
                                if (len > line.Length - idx + offset) len = (int)(line.Length - idx + offset);
                                if (len < 0) len = (int)(line.Length - idx + offset);
                                runText = line.Substring((int)(idx - offset), len);
                            }
                            else
                            {
                                runText = line.Substring((int)(idx - offset));
                            }                            

                            //if (r != null)
                            //if ((idx - offset) + r.Length < line.Length)
                            //{
                            //    runText = line.Substring((int)(idx - offset), (int)r.Length);
                            //}                            

                            _writer.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingML);
                            if (r != null)
                            {
                                string dummy = "";
                                new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", textbox.FirstAncestorWithType<Slide>(),ref dummy);
                            }
                            else
                            {
                                //textbox.FirstAncestorWithType<Slide>()
                            }

                            _writer.WriteStartElement("a", "t", OpenXmlNamespaces.DrawingML);
                            _writer.WriteValue(runText);
                            _writer.WriteEndElement();

                            _writer.WriteEndElement();

                            idx += (uint)runText.Length; // +1;

                        }

                        first = false;

                       
                    }

                    _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();

                    idx += 1;
                }

            }

        }

        private void writeP(ParagraphRun p)
        {
            _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

            _writer.WriteStartElement("a", "pPr", OpenXmlNamespaces.DrawingML);
            if (!(p == null))
            {
                if (p.IndentLevel > 0) _writer.WriteAttributeString("lvl", p.IndentLevel.ToString());
                if (p.LeftMarginPresent) _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)p.LeftMargin).ToString());
                if (p.IndentPresent) _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(p.LeftMargin - p.Indent)))).ToString());

                if (p.AlignmentPresent)
                {
                    switch (p.Alignment)
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

                if (p.LineSpacingPresent)
                {
                    _writer.WriteStartElement("a", "lnSpc", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteAttributeString("val", (p.LineSpacing * 1000).ToString());
                    //_writer.WriteEndElement(); //spcPct

                    if (p.LineSpacing < 0)
                    {
                        _writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (-1 * p.LineSpacing * 12).ToString()); //TODO: this has to be verified!
                        _writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (1000 * p.LineSpacing).ToString());
                        _writer.WriteEndElement(); //spcPct
                    }


                    _writer.WriteEndElement(); //lnSpc
                }
                if (p.SpaceBeforePresent)
                {
                    _writer.WriteStartElement("a", "spcBef", OpenXmlNamespaces.DrawingML);
                    if (p.SpaceBefore < 0)
                    {
                        _writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (-1 * 12 * p.SpaceBefore).ToString()); //TODO: the 12 is wrong: find correct value
                        _writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (1000 * p.SpaceBefore).ToString());
                        _writer.WriteEndElement(); //spcPct
                    }
                    _writer.WriteEndElement(); //spcBef
                }
                
                if (p.SpaceAfterPresent)
                {
                    _writer.WriteStartElement("a", "spcAft", OpenXmlNamespaces.DrawingML);
                    if (p.SpaceAfter < 0)
                    {
                        _writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (-1 * 12 * p.SpaceAfter).ToString()); //TODO: the 12 is wrong: find correct value
                        _writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (1000 * p.SpaceAfter).ToString());
                        _writer.WriteEndElement(); //spcPct
                    }
                    _writer.WriteEndElement(); //spcAft
                }

                if (p.BulletFlagsFieldPresent)
                {
                    if ((p.BulletFlags & (ushort)ParagraphMask.HasBullet) == 0)
                    {
                        _writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");
                    }
                    else
                    {
                        if (p.BulletSizePresent)
                        {
                            if (p.BulletSize > 0)
                            {
                                _writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("val", (p.BulletSize * 1000).ToString());
                                _writer.WriteEndElement(); //buChar
                            }
                            else
                            {
                                //TODO
                            }
                        }
                        if (p.BulletFontPresent)
                        {                           
                            _writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                            FontCollection fonts = _ctx.Ppt.DocumentRecord.FirstChildWithType<DIaLOGIKa.b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                            FontEntityAtom entity = fonts.entities[(int)p.BulletTypefaceIdx];
                            if (entity.TypeFace.IndexOf('\0') > 0)
                            {
                                _writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                            }
                            else
                            {
                                _writer.WriteAttributeString("typeface", entity.TypeFace);
                            }
                            _writer.WriteEndElement(); //buChar
                        }
                        if (p.BulletCharPresent)
                        {
                            _writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("char", p.BulletChar.ToString());
                            _writer.WriteEndElement(); //buChar
                        }
                    }
                }

            }
            _writer.WriteEndElement(); //pPr
        }
    }
}
