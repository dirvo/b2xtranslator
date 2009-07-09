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
        protected static ParagraphRun GetParagraphRun(TextStyleAtom style, uint forIdx, ref int runCount)
        {
            if (style == null)
                return null;

            uint idx = 0;
            runCount = 0;

            foreach (ParagraphRun p in style.PRuns)
            {
                if (forIdx < idx + p.Length)
                    return p;

                idx += p.Length;
                runCount++;
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
            Record rec = Record.ReadRecord(ms);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            FooterMCAtom mca = null;
            TextRulerAtom ruler = null;
            List<MouseClickInteractiveInfoContainer> mciics = new List<MouseClickInteractiveInfoContainer>();
            MasterTextPropAtom masterTextProp = null;
            string text = "";
            string origText = "";
            ShapeOptions so = textbox.FirstAncestorWithType<ShapeContainer>().FirstChildWithType<ShapeOptions>();
            TextMasterStyleAtom defaultStyle = null;
            int lvl = 0;

            switch (rec.TypeCode)
            {
                case 3999:
                    thAtom = (TextHeaderAtom)rec;
                    while (ms.Position < ms.Length)
                    {
                        rec = Record.ReadRecord(ms);

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
                            case 0xff2: //MouseClickInteractiveInfoContainer
                                mciics.Add((MouseClickInteractiveInfoContainer)rec);
                                break;
                            case 0xfdf: //MouseClickTextInteractiveInfoAtom
                                mciics[mciics.Count-1].Range = (MouseClickTextInteractiveInfoAtom)rec;
                                break;
                            case 0xff7: //DateTimeMCAtom
                                DateTimeMCAtom d = (DateTimeMCAtom)rec;
                                String date = System.DateTime.Now.ToString();

                                //_writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

                                int runCount = 0;
                                ParagraphRun p = GetParagraphRun(style, 0, ref runCount);
                                MasterTextPropRun tp = GetMasterTextPropRun(masterTextProp, 0);
                                writeP(p, tp, so, ruler, defaultStyle,0);

                                _writer.WriteStartElement("a", "fld", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("id", "{1023E2E8-AA53-4FEA-8F5C-1FABD68F61AB}");
                                _writer.WriteAttributeString("type", "datetime1");

                                CharacterRun r = GetCharacterRun(style, 0);
                                if (r != null)
                                {
                                    string dummy = "";
                                    string dummy2 = "";
                                    string dummy3 = "";
                                    RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                                    if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                                    if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                                    new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", slide, ref dummy, ref dummy2, ref dummy3, lang, defaultStyle,lvl,mciics,pparentShapeTreeMapping,0);
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
                        rec = Record.ReadRecord(ms);
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
                    if (Tools.Utils.BitmaskToBool(a.Flags, 0x01 << 1) && a.MasterId > 0)
                    {
                        Slide m = _ctx.Ppt.FindMasterRecordById(a.MasterId);
                        foreach (TextMasterStyleAtom at in m.AllChildrenWithType<TextMasterStyleAtom>())
                        {
                            //if (at.Instance == 1 && thAtom.TextType == TextType.Other)
                            //{
                            //    defaultStyle = at;
                            //    break;
                            //}
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

                    int runCount = 0;
                    ParagraphRun p = GetParagraphRun(style, idx, ref runCount);
                    MasterTextPropRun tp = GetMasterTextPropRun(masterTextProp, idx);

                    if (p != null) lvl = p.IndentLevel;

                    String runText;

                    if (runlines.Length > 0)
                    if (runlines[0].StartsWith("\t"))
                    {

                    }

                    writeP(p, tp, so, ruler, defaultStyle, runCount);

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
                                new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", slide, ref dummy, ref dummy2, ref dummy3, lang, defaultStyle,lvl,mciics,pparentShapeTreeMapping,idx);
                            }

                            _writer.WriteEndElement();
                            if (line.Length == 0)
                            {
                               idx++;
                               internalOffset -= 1;
                            }
                            
                            internalOffset += 1;
                            
                        }

                        

                        while (idx < offset + line.Length)
                        {
                            textwritten = true;
                            len = line.Length;
                            CharacterRun r = null;
                            if (idx + (uint)internalOffset == 0)
                            {
                                r = GetCharacterRun(style, 0);
                                CharacterRunStart = GetCharacterRunStart(style, 0);
                            }
                            else
                            {
                                r = GetCharacterRun(style, idx + (uint)internalOffset);
                                CharacterRunStart = GetCharacterRunStart(style, idx + (uint)internalOffset);
                            }
                            if (r != null)
                            {
                                
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

                            string dummy = "";
                            string dummy2 = "";
                            string dummy3 = "";
                            RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();

                            if (r != null || defaultStyle != null)
                            {
                                new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", slide,ref dummy, ref dummy2, ref dummy3, lang, defaultStyle,lvl, mciics, pparentShapeTreeMapping, idx);
                            }
                            else
                            {                              
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
                            new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "endParaRPr", slide, ref dummy, ref dummy2, ref dummy3, lang, defaultStyle,lvl,mciics,pparentShapeTreeMapping,idx);
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

        private void writeP(ParagraphRun p, MasterTextPropRun tp, ShapeOptions so, TextRulerAtom ruler, TextMasterStyleAtom defaultStyle, int runCount)
        {
            int writtenLeftMargin = -1;
            _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

            _writer.WriteStartElement("a", "pPr", OpenXmlNamespaces.DrawingML);
            if (p == null)
            {
                _writer.WriteAttributeString("lvl", tp.indentLevel.ToString());
                               

                if (defaultStyle != null && defaultStyle.PRuns.Count > tp.indentLevel)
                {
                    if (defaultStyle.PRuns[tp.indentLevel].LeftMarginPresent)
                    {
                        _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)defaultStyle.PRuns[tp.indentLevel].LeftMargin).ToString());
                        writtenLeftMargin = (int)defaultStyle.PRuns[tp.indentLevel].LeftMargin;
                    }
                    if (defaultStyle.PRuns[tp.indentLevel].IndentPresent)
                    {
                        _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(defaultStyle.PRuns[tp.indentLevel].LeftMargin - defaultStyle.PRuns[tp.indentLevel].Indent)))).ToString());
                    }


                    if (defaultStyle.PRuns[tp.indentLevel].AlignmentPresent)
                    {
                        switch (defaultStyle.PRuns[tp.indentLevel].Alignment)
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

                    if (defaultStyle.PRuns[tp.indentLevel].BulletFlagsFieldPresent)
                    {
                        if ((defaultStyle.PRuns[tp.indentLevel].BulletFlags & (ushort)ParagraphMask.HasBullet) == 0)
                        {
                            _writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");
                        }
                        else
                        {
                            //if (defaultStyle.PRuns[tp.indentLevel].BulletColorPresent)
                            //{
                            //    _writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                            //    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            //    string s = defaultStyle.PRuns[tp.indentLevel].BulletColor.Red.ToString("X").PadLeft(2, '0') + defaultStyle.PRuns[tp.indentLevel].BulletColor.Green.ToString("X").PadLeft(2, '0') + defaultStyle.PRuns[tp.indentLevel].BulletColor.Blue.ToString("X").PadLeft(2, '0');
                            //    _writer.WriteAttributeString("val", s);
                            //    _writer.WriteEndElement();
                            //    _writer.WriteEndElement(); //buClr
                            //}
                            if (defaultStyle.PRuns[tp.indentLevel].BulletSizePresent)
                            {
                                if (defaultStyle.PRuns[tp.indentLevel].BulletSize > 0)
                                {
                                    _writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", (defaultStyle.PRuns[tp.indentLevel].BulletSize * 1000).ToString());
                                    _writer.WriteEndElement(); //buSzPct
                                }
                                else
                                {
                                    //TODO
                                }
                            }
                            if (defaultStyle.PRuns[tp.indentLevel].BulletFontPresent)
                            {
                                _writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                                FontCollection fonts = _ctx.Ppt.DocumentRecord.FirstChildWithType<DIaLOGIKa.b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                                FontEntityAtom entity = fonts.entities[(int)defaultStyle.PRuns[tp.indentLevel].BulletTypefaceIdx];
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


                            if (parentShapeTreeMapping != null && parentShapeTreeMapping.ShapeStyleTextProp9Atom != null && parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs.Count > runCount && parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].fBulletHasAutoNumber == 1 && parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].bulletAutoNumberScheme == -1)
                            {
                                _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("type", "arabicPeriod");
                                _writer.WriteEndElement();
                            }
                            else if (defaultStyle.PRuns[tp.indentLevel].BulletCharPresent)
                            {
                                _writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("char", defaultStyle.PRuns[tp.indentLevel].BulletChar.ToString());
                                _writer.WriteEndElement(); //buChar
                            }
                            else if (!defaultStyle.PRuns[tp.indentLevel].BulletCharPresent)
                            {
                                _writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("char", "�");
                                _writer.WriteEndElement(); //buChar
                            }

                        }
                    }
                }
            }
            else
            {
                if (p.IndentLevel > 0) _writer.WriteAttributeString("lvl", p.IndentLevel.ToString());
                if (p.LeftMarginPresent)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)p.LeftMargin).ToString());
                    writtenLeftMargin = (int)p.LeftMargin;
                } 
                else if (ruler != null && ruler.fLeftMargin1 && p.IndentLevel == 0){
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin1).ToString());
                    writtenLeftMargin = ruler.leftMargin1;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent1 && p.IndentLevel == 0)))
                    {
                        _writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin2 && p.IndentLevel == 1)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin2).ToString());
                    writtenLeftMargin = ruler.leftMargin2;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent2 && p.IndentLevel == 1)))
                    {
                        _writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin3 && p.IndentLevel == 2)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin3).ToString());
                    writtenLeftMargin = ruler.leftMargin3;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent3 && p.IndentLevel == 2)))
                    {
                        _writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin4 && p.IndentLevel == 3)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin4).ToString());
                    writtenLeftMargin = ruler.leftMargin4;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent4 && p.IndentLevel == 3)))
                    {
                        _writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin5 && p.IndentLevel == 4)
                {
                    _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin5).ToString());
                    writtenLeftMargin = ruler.leftMargin5;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent5 && p.IndentLevel == 4)))
                    {
                        _writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextLeft) && so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.TextBooleanProperties))
                {
                    TextBooleanProperties props = new TextBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.TextBooleanProperties].op);
                    if (props.fUsefAutoTextMargin && (props.fAutoTextMargin == false))
                    if (so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op > 0)
                    _writer.WriteAttributeString("marL", so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op.ToString());
                    //writtenLeftMargin = Utils.EMUToMasterCoord((int)so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op);
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
                    if (writtenLeftMargin == -1 )
                    {
                        writtenLeftMargin = (int)(defaultStyle.PRuns[p.IndentLevel].LeftMargin);
                    }
                    _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(writtenLeftMargin - defaultStyle.PRuns[p.IndentLevel].Indent)))).ToString());
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
                                                
                        if (p.BulletColorPresent)
                        {
                            _writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                            
                            string s = p.BulletColor.Red.ToString("X").PadLeft(2, '0') + p.BulletColor.Green.ToString("X").PadLeft(2, '0') + p.BulletColor.Blue.ToString("X").PadLeft(2, '0');
                            switch (p.BulletColor.Index)
                            {
                                case 7:
                                    _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", "folHlink");
                                    _writer.WriteEndElement();
                                    break;
                                case 6:
                                    _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", "hlink");
                                    _writer.WriteEndElement();
                                    break;
                                case 3:
                                    _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", "tx2");
                                    _writer.WriteEndElement();
                                    break;
                                case 2:
                                    _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", "bg2");
                                    _writer.WriteEndElement();
                                    break;
                                case 1:
                                    _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", "bg1");
                                    _writer.WriteEndElement();
                                    break;
                                default:
                                    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("val", s);
                                    _writer.WriteEndElement();
                                    break;
                            }
                           
                            _writer.WriteEndElement(); //buClr
                        }
                        if (p.BulletSizePresent)
                        {
                            if (p.BulletSize > 0)
                            {
                                _writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("val", (p.BulletSize * 1000).ToString());
                                _writer.WriteEndElement(); //buSzPct
                            }
                            else
                            {
                                //TODO
                            }
                        }
                        else if (p.BulletFlagsFieldPresent && (p.BulletFlags & 0x1 << 3) > 0)
                        {
                            _writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", "75000");
                            _writer.WriteEndElement(); //buSzPct
                        }

                        if (p.BulletFontPresent)
                        {
                            if (!(p.BulletFlagsFieldPresent && (p.BulletFlags & 0x1 << 1) == 0))
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
                            else
                            {
                                _writer.WriteElementString("a", "buFontTx", OpenXmlNamespaces.DrawingML, "");
                            }
                        }

                        bool autoNumberingWritten = false;
                        if (_ctx.Ppt.DocumentRecord.DocInfoListContainer.FirstDescendantWithType<OutlineTextProps9Container>() != null)
                        {
                            OutlineTextProps9Container c = _ctx.Ppt.DocumentRecord.DocInfoListContainer.FirstDescendantWithType<OutlineTextProps9Container>();
                            Slide slide = so.FirstAncestorWithType<Slide>();

                            if (slide != null)
                            foreach (OutlineTextProps9Entry entry in c.OutlineTextProps9Entries)
                            {
                                if (slide.PersistAtom.SlideId == entry.outlineTextHeaderAtom.slideIdRef)
                                {
                                    if (entry.styleTextProp9Atom.P9Runs.Count > runCount && entry.styleTextProp9Atom.P9Runs[runCount].fBulletHasAutoNumber == 1)
                                    {
                                        switch (entry.styleTextProp9Atom.P9Runs[runCount].bulletAutoNumberScheme)
                                        {
                                            case -1:
                                            case 3:
                                                _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                                _writer.WriteAttributeString("type", "arabicPeriod");
                                                if (entry.styleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                                {
                                                    _writer.WriteAttributeString("startAt", entry.styleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                                }
                                                _writer.WriteEndElement();
                                                autoNumberingWritten = true;
                                                break;
                                            case 1:
                                                _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                                _writer.WriteAttributeString("type", "alphaUcPeriod");
                                                if (entry.styleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                                {
                                                    _writer.WriteAttributeString("startAt", entry.styleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                                }
                                                _writer.WriteEndElement();
                                                autoNumberingWritten = true;
                                                break;
                                        }
                                    }
                                }
                            }

                            //OutlineTextPropsHeader9Atom a = c.FirstChildWithType<OutlineTextPropsHeader9Atom>();
                            //Slide slide = so.FirstAncestorWithType<Slide>();
                            //if (slide.PersistAtom.SlideId == a.slideIdRef)
                            //{
                            //    StyleTextProp9Atom s = c.FirstChildWithType<StyleTextProp9Atom>();
                            //    if (s.P9Runs.Count > runCount && s.P9Runs[runCount].fBulletHasAutoNumber == 1)
                            //    {
                            //        switch (s.P9Runs[runCount].bulletAutoNumberScheme)
                            //        {
                            //            case -1:
                            //                _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                            //                _writer.WriteAttributeString("type", "arabicPeriod");
                            //                _writer.WriteEndElement();
                            //                autoNumberingWritten = true;
                            //                break;
                            //            case 1:
                            //                _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                            //                _writer.WriteAttributeString("type", "alphaUcPeriod");
                            //                _writer.WriteEndElement();
                            //                autoNumberingWritten = true;
                            //                break;
                            //        }
                            //    }
                            //}
                        }

                        if (!autoNumberingWritten)
                        {
                            if (parentShapeTreeMapping != null && parentShapeTreeMapping.ShapeStyleTextProp9Atom != null && parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs.Count > runCount && parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].fBulletHasAutoNumber == 1)
                            {
                                switch(parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].bulletAutoNumberScheme)
                                {
                                    case -1:
                                    case 3:
                                        _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                        _writer.WriteAttributeString("type", "arabicPeriod");
                                        if (parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                        {
                                            _writer.WriteAttributeString("startAt", parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                        }
                                        _writer.WriteEndElement();
                                        break;
                                    case 1:
                                        _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                        _writer.WriteAttributeString("type", "alphaUcPeriod");
                                        if (parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                        {
                                            _writer.WriteAttributeString("startAt", parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                        }
                                        _writer.WriteEndElement();
                                        break;
                                }
                            }
                            else if (p.BulletCharPresent)
                            {
                                _writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("char", p.BulletChar.ToString());
                                _writer.WriteEndElement(); //buChar

                                Slide s = so.FirstAncestorWithType<Slide>();
                            }
                            else if (!p.BulletCharPresent)
                            {
                                _writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                _writer.WriteAttributeString("char", "�");
                                _writer.WriteEndElement(); //buChar
                            }
                        }
                    }
                }
            }
            
            _writer.WriteEndElement(); //pPr
        }
                
    }
}
