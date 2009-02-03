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

        public void Apply(ClientTextbox textbox)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream(textbox.Bytes);
            Record rec = Record.ReadRecord(ms, 0);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
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
                            default:
                                TextAtom textAtom = thAtom.TextAtom;
                                text = (textAtom == null) ? "" : textAtom.Text;
                                style = thAtom.TextStyleAtom;
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
                String[] runlines = text.Split(new char[] { '\v', '\r' });
                foreach (String line in runlines)
                {
                    //each line forms a paragraph
                    ParagraphRun p = GetParagraphRun(style, idx);
                    String runText;
                    writeP(p);
                    uint offset = idx;

                    while (idx < offset + line.Length)
                    {
                        CharacterRun r = GetCharacterRun(style, idx);

                        runText = line.Substring((int)(idx-offset));

                        if (r != null)
                            if ((idx - offset) + r.Length < line.Length)
                        {
                            runText = line.Substring((int)(idx - offset), (int)r.Length);
                        }
                        

                        _writer.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingML);
                        if (r != null)
                        {
                            new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", textbox.FirstAncestorWithType<Slide>());
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

                    _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();

                    idx += 1;
                }

            }


            //TraceLogger.DebugInternal("TextMapping: text = {0}", Tools.Utils.StringInspect(text));            

            //// Special case: always write out at least one paragraph (even if idx == text.Length == 0)
            //while (idx < text.Length || text.Length == 0)
            //{
            //    ParagraphRun p = GetParagraphRun(style, idx);

            //    uint pEndIdx = (p != null) ? (uint)Math.Min(idx + p.Length, text.Length) : (uint)text.Length;

            //    TraceLogger.DebugInternal("Paragraph run from {0} to {1}", idx, pEndIdx);

            //    if (idx == pEndIdx)
            //    {
            //        _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);
            //        _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
            //        // TODO...
            //        _writer.WriteEndElement();
            //        _writer.WriteEndElement();
            //    }

            //    bool startNewParagraph = true;
            //    writeP(p);

            //    while (idx < pEndIdx)
            //    {
            //        CharacterRun r = GetCharacterRun(style, idx);
            //        // Current run length or remaining text length if no run available
            //        uint rlen = (r != null) ? r.Length : (uint)(text.Length - idx);

            //        // Remaining paragraph length
            //        uint plen = pEndIdx - idx;

            //        // Remaining text length
            //        uint tlen = (uint)(text.Length - idx);

            //        // Length of extracted runText can't go beyond character run,
            //        // remaining paragraph run and remaining text length so limit it.
            //        uint slen = rlen;
            //        if (slen > tlen)
            //            slen = tlen;
            //        if (slen > plen)
            //            slen = plen;

            //        String runText = text.Substring((int)idx, (int)slen);
            //        //bool isLastRunOfParagraph = idx + slen == pEndIdx;
            //        //if (isLastRunOfParagraph)
            //            //runText = runText.TrimEnd(new char[] { '\v', '\r', '\n' });
            //        if (runText.EndsWith("\r") | runText.EndsWith("\v") | runText.EndsWith("\n"))
            //        {
            //            //runText = runText.Substring(0, runText.Length - 1);
            //            //startNewParagraph = true;
            //        }

            //        TraceLogger.DebugInternal("Character run from {0} to {1} ({3}): {2}",
            //            idx, idx + rlen, Tools.Utils.StringInspect(runText), slen);

            //        String[] lines = new String[1];
            //        if (runText.Length > 1) { lines = runText.Split(new char[] { '\v', '\r' }); } else { lines[0] = runText.TrimEnd(new char[] { '\v', '\r' }); };

            //        //bool isFirstLine = true;
            //        int lineIdx = 0;

            //        TraceLogger.DebugInternal("Split runtext {0} into these lines:", Tools.Utils.StringInspect(runText));

            //        foreach (String line in lines)
            //        {                        
            //            TraceLogger.DebugInternal("  {0}", Tools.Utils.StringInspect(line));

            //            if (startNewParagraph)
            //            {
            //                _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
            //                _writer.WriteEndElement();
            //                _writer.WriteEndElement();      

            //                p = GetParagraphRun(style, idx);
            //                writeP(p);
            //            }

            //            if (line.Length > 0)
            //            {
            //                _writer.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingML);
            //                if (r != null) new CharacterRunPropsMapping(_ctx, _writer).Apply(r, "rPr", textbox.FirstAncestorWithType<Slide>());

            //                _writer.WriteStartElement("a", "t", OpenXmlNamespaces.DrawingML);
            //                _writer.WriteValue(line);
            //                _writer.WriteEndElement();

            //                _writer.WriteEndElement();
                                                        
            //            }

            //            lineIdx += line.Length + 1;
            //            //isFirstLine = false;
                        
            //        }

            //        idx += rlen;
            //    }

            //    _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
            //    _writer.WriteEndElement();
            //    _writer.WriteEndElement();               

            //    idx = pEndIdx;

            //    /* Didn't move so stop looping */
            //    if (text.Length == 0)
            //        break;
            //}
        }

        private void writeP(ParagraphRun p)
        {
            _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

            _writer.WriteStartElement("a", "pPr", OpenXmlNamespaces.DrawingML);
            if (!(p == null))
            {
                if (p.IndentLevel > 0) _writer.WriteAttributeString("lvl", p.IndentLevel.ToString());
                if (p.IndentPresent) _writer.WriteAttributeString("indent", p.Indent.ToString());
                if (p.BulletCharPresent)
                {
                   
                    //_writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteAttributeString("char", p.BulletChar.ToString());
                    //_writer.WriteEndElement(); //buChar
                }
                if (p.LineSpacingPresent)
                {
                    _writer.WriteStartElement("a", "lnSpc", OpenXmlNamespaces.DrawingML);
                    _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", (p.LineSpacing * 1000).ToString());
                    _writer.WriteEndElement(); //spcPct
                    _writer.WriteEndElement(); //lnSpc
                }

            }
            _writer.WriteEndElement(); //pPr
        }
    }
}
