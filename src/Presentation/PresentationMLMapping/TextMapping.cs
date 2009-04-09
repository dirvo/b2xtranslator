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
        private string lang = "en-US";
        private ShapeTreeMapping parentShapeTreeMapping = null;

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
        /// Returns the MasterTextParagraphRun of the given style that is active at the given index.
        /// </summary>
        /// <param name="style">style to use</param>
        /// <param name="forIdx">index to use</param>
        /// <returns>ParagraphRun or null in case no run was found</returns>
        protected static MasterTextPropRun GetMasterTextPropRun(MasterTextPropAtom style, uint forIdx)
        {
            if (style == null)
                return new MasterTextPropRun();

            uint idx = 0;

            foreach (MasterTextPropRun p in style.MasterTextPropRuns)
            {
                if (forIdx < idx + p.count)
                    return p;

                idx += p.count;
            }

            return new MasterTextPropRun();
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
            Apply(null, textbox, "", "", "");
        }

        public void Apply(ShapeTreeMapping pparentShapeTreeMapping, ClientTextbox textbox, string footertext, string headertext, string datetext)
        {
            parentShapeTreeMapping = pparentShapeTreeMapping;
            System.IO.MemoryStream ms = new System.IO.MemoryStream(textbox.Bytes);
            Record rec = Record.ReadRecord(ms, 0);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            FooterMCAtom mca = null;
            TextRulerAtom ruler = null;
            MasterTextPropAtom masterTextProp = null;
            string text = "";
            string origText = "";
            ShapeOptions so = textbox.FirstAncestorWithType<ShapeContainer>().FirstChildWithType<ShapeOptions>();
            TextMasterStyleAtom defaultStyle = null;

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
                                origText = text;
                                thAtom.TextAtom = (TextAtom)rec;
                                break;
                            case 0xfa1: //TextRunStyleAtom
                                style = (TextStyleAtom)rec;
                                style.TextHeaderAtom = thAtom;
                                break;
                            case 0xfa6: //TextRulerAtom
                                ruler = (TextRulerAtom)rec;
                                break;
                            case 0xfa8: //TextBytesAtom
                                text = ((TextBytesAtom)rec).Text;
                                origText = text;
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

                                text = text.Replace(origText.Substring(snmca.Position, 1), "");

                                break;
                            case 0xff7: //DateTimeMCAtom
                                DateTimeMCAtom d = (DateTimeMCAtom)rec;
                                String date = System.DateTime.Now.ToString();

                                //_writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

                                ParagraphRun p = GetParagraphRun(style, 0);
                                MasterTextPropRun tp = GetMasterTextPropRun(masterTextProp, 0);
                                writeP(p, tp, so, ruler, defaultStyle);

                                _writer.WriteStartElement("a", "fld", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("id", "{1023E2E8-AA53-4FEA-8F5C-1FABD68F61AB}");
                                _writer.WriteAttributeString("type", "datetime9");

                                CharacterRun r = GetCharacterRun(style, 0);
                                if (r != null)
                                {
                                    string dummy = "";
                                    string dummy2 = "";
                                    string dummy3 = "";
                                    RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                                    if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                                    if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                                    new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", slide, ref dummy, ref dummy2, ref dummy3, lang, defaultStyle);
                                }

                                _writer.WriteElementString("a", "t", OpenXmlNamespaces.DrawingML, date);
                                _writer.WriteEndElement(); //fld                                
                                _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                                _writer.WriteEndElement(); //endParaRPr
                                _writer.WriteEndElement(); //p

                                text = text.Replace(origText.Substring(d.Position, 1), "");

                                foreach (CharacterRun run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            case 0xff9: //HeaderMCAtom
                                HeaderMCAtom hmca = (HeaderMCAtom)rec;
                                text = text.Replace(origText.Substring(hmca.Position, 1), headertext);

                                foreach (CharacterRun run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            case 0xffa: //FooterMCAtom
                                mca = (FooterMCAtom)rec;
                                text = text.Replace(origText.Substring(mca.Position, 1), footertext);

                                foreach (CharacterRun run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            case 0xff8: //GenericDateMCAtom
                                GenericDateMCAtom gdmca = (GenericDateMCAtom)rec;
                                text = text.Replace(origText.Substring(gdmca.Position, 1), datetext);

                                foreach (CharacterRun run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
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

                    while (ms.Position < ms.Length)
                    {
                        rec = Record.ReadRecord(ms, 0);
                        switch (rec.TypeCode)
                        {
                            case 0xfa6: //TextRulerAtom
                                ruler = (TextRulerAtom)rec;
                                break;
                            default:
                                break;
                        }
                    }

                    break;
                default:
                    throw new NotSupportedException("Can't find text for ClientTextbox without TextHeaderAtom and OutlineTextRefAtom");
            }

            uint idx = 0;                      

            Slide s = textbox.FirstAncestorWithType<Slide>();
           
            if (s != null)
            {
                try
                {
                    SlideAtom a = s.FirstChildWithType<SlideAtom>();
                    if (a.MasterId > 0)
                    {
                        Slide m = _ctx.Ppt.FindMasterRecordById(a.MasterId);
                        foreach (TextMasterStyleAtom at in m.AllChildrenWithType<TextMasterStyleAtom>())
                        {
                            if (at.Instance == 1 && thAtom.TextType == TextType.Other)
                            {
                                //defaultStyle = at;
                                break;
                            }
                            if (at.Instance == (int)thAtom.TextType)
                            {
                                defaultStyle = at;
                            }
                        }
                    }
                }
                catch (Exception)
                {                    
                    throw;
                }
                
            }
            
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.hspMaster))
            {
                uint MasterID = so.OptionsByID[ShapeOptions.PropertyId.hspMaster].op;
            }

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
                    MasterTextPropRun tp = GetMasterTextPropRun(masterTextProp, idx);

                    String runText;

                    if (runlines.Length > 0)
                    if (runlines[0].StartsWith("\t"))
                    {

                    }

                    writeP(p, tp, so, ruler, defaultStyle);

                    uint CharacterRunStart;
                    int len;

                    bool first = true;
                    bool textwritten = false;
                    foreach (string line in runlines)
                    {
                        uint offset = idx;

                        if (!first)
                        {
                            _writer.WriteStartElement("a", "br", OpenXmlNamespaces.DrawingML);
                            CharacterRun r = GetCharacterRun(style, idx + (uint)internalOffset + 1);
                            if (r != null)
                            {
                                string dummy = "";
                                string dummy2 = "";
                                string dummy3 = "";
                                RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                                if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                                if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                                new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", slide, ref dummy, ref dummy2, ref dummy3, lang, defaultStyle);
                            }

                            _writer.WriteEndElement();
                            internalOffset += 1;
                        }

                        while (idx < offset + line.Length)
                        {
                            textwritten = true;
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
                                string dummy2 = "";
                                string dummy3 = "";
                                RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                                if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                                if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                                new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", slide,ref dummy, ref dummy2, ref dummy3, lang, defaultStyle);
                            }
                            else
                            {
                                //TODO: read real language
                                _writer.WriteStartElement("a", "rPr", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("lang", lang);
                                _writer.WriteEndElement();
                            }

                            _writer.WriteStartElement("a", "t", OpenXmlNamespaces.DrawingML);

                            _writer.WriteValue(runText.Replace(Char.ToString((char)0x05), ""));
                            
                            _writer.WriteEndElement();

                            _writer.WriteEndElement();

                            idx += (uint)runText.Length; // +1;

                        }

                        first = false;

                       
                    }

                    if (!textwritten)
                    {
                        CharacterRun r = GetCharacterRun(style, idx + (uint)internalOffset);
                        if (r != null)
                        {
                            string dummy = "";
                            string dummy2 = "";
                            string dummy3 = "";
                            RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                            new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "endParaRPr", slide, ref dummy, ref dummy2, ref dummy3, lang, defaultStyle);
                        }

                    }
                    else
                    {
                        _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                        _writer.WriteEndElement();
                    }

                    _writer.WriteEndElement();

                    idx += 1;
                }

            }

        }

        private void writeP(ParagraphRun p, MasterTextPropRun tp, ShapeOptions so, TextRulerAtom ruler, TextMasterStyleAtom defaultStyle)
        {
            _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

            _writer.WriteStartElement("a", "pPr", OpenXmlNamespaces.DrawingML);
            if (p == null)
            {
                _writer.WriteAttributeString("lvl", tp.indentLevel.ToString()); 
            }
            else
            {
                if (p.IndentLevel > 0) _writer.WriteAttributeString("lvl", p.IndentLevel.ToString());
                if (p.LeftMarginPresent)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)p.LeftMargin).ToString());
                } 
                else if (ruler != null && ruler.fLeftMargin1 && p.IndentLevel == 0){
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin1).ToString());
                }
                else if (ruler != null && ruler.fLeftMargin2 && p.IndentLevel == 1)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin2).ToString());
                }
                else if (ruler != null && ruler.fLeftMargin3 && p.IndentLevel == 2)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin3).ToString());
                }
                else if (ruler != null && ruler.fLeftMargin4 && p.IndentLevel == 3)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin4).ToString());
                }
                else if (ruler != null && ruler.fLeftMargin5 && p.IndentLevel == 4)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin5).ToString());
                } else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextLeft))
                {
                    TextBooleanProperties props = new TextBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.TextBooleanProperties].op);
                    if (props.fUsefAutoTextMargin && (props.fAutoTextMargin == false))
                    if (so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op > 0)
                    _writer.WriteAttributeString("marL", so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op.ToString());
                }
                if (p.IndentPresent)
                {
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(p.LeftMargin - p.Indent)))).ToString());
                }
                else if (ruler != null && ruler.fIndent1 && p.IndentLevel == 0)
                {
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin1 - ruler.indent1)))).ToString());
                }
                else if (ruler != null && ruler.fIndent2 && p.IndentLevel == 1)
                {
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin2 - ruler.indent2)))).ToString());
                }
                else if (ruler != null && ruler.fIndent3 && p.IndentLevel == 2)
                {
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin3 - ruler.indent3)))).ToString());
                }
                else if (ruler != null && ruler.fIndent4 && p.IndentLevel == 3)
                {
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin4 - ruler.indent4)))).ToString());
                }
                else if (ruler != null && ruler.fIndent5 && p.IndentLevel == 4)
                {
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin5 - ruler.indent5)))).ToString());
                }
                else if (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent)
                {
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(defaultStyle.PRuns[p.IndentLevel].LeftMargin - defaultStyle.PRuns[p.IndentLevel].Indent)))).ToString());
                }
              
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
                else if (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel)
                {
                    if (defaultStyle.PRuns[p.IndentLevel].AlignmentPresent)
                    {
                        switch (defaultStyle.PRuns[p.IndentLevel].Alignment)
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
                }

                if (p.DefaultTabSizePresent)
                {
                    _writer.WriteAttributeString("defTabSz", Utils.MasterCoordToEMU((int)p.DefaultTabSize).ToString());
                }
                else if (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel)
                {
                    if (defaultStyle.PRuns[p.IndentLevel].DefaultTabSizePresent)
                    {
                        _writer.WriteAttributeString("defTabSz", Utils.MasterCoordToEMU((int)defaultStyle.PRuns[p.IndentLevel].DefaultTabSize).ToString());
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

                        if (parentShapeTreeMapping != null && parentShapeTreeMapping.ShapeStyleTextProp9Atom != null && parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs.Count > p.IndentLevel)
                        {
                            if (parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[p.IndentLevel].fBulletHasAutoNumber == 1)
                            {
                                if (parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[p.IndentLevel].bulletAutoNumberScheme != -1)
                                {

                                }
                                else
                                {
                                    _writer.WriteStartElement("a","buAutoNum",OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("type","arabicPeriod");
                                    _writer.WriteEndElement();
                                }
                            }
                           
                        }

                    }
                }

            }
            _writer.WriteEndElement(); //pPr
        }
    }
}
