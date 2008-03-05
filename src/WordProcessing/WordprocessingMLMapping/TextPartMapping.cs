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
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TextPartMapping : 
        AbstractOpenXmlMapping,
        IMapping<WordDocument>
    {

        public TextPartMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(WordDocument visited)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "document", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "body", OpenXmlNamespaces.WordprocessingML);

            bool suppressNextChar = false;

            //walk through the characters
            for(int i=0; i<visited.Text.Count; i++)
            {
                if (i == 0)
                {
                    startNewParagraph(i, visited);
                }

                char c = visited.Text[i];

                //check the char
                if(c == TextBoundary.AnnotationReference)
                {
                }
                else if (c == TextBoundary.AutoNumberedFootnoteReference)
                {
                }
                else if (c == TextBoundary.BreakingHyphen)
                {
                }
                else if (c == TextBoundary.CellOrRowMark)
                {
                }
                else if (c == TextBoundary.ColumnBreak)
                {
                }
                else if (c == TextBoundary.CurrentPageNumber)
                {
                }
                else if (c == TextBoundary.FieldBeginMark)
                {
                    suppressNextChar = true;
                }
                else if (c == TextBoundary.FieldEndMark)
                {
                    suppressNextChar = false;
                }
                else if (c == TextBoundary.FieldSeperator)
                {
                }
                else if (c == TextBoundary.FootnoteContinuation)
                {
                }
                else if (c == TextBoundary.FootnoteSeparator)
                {
                }
                else if (c == TextBoundary.HandAnnotationPicture)
                {
                }
                else if (c == TextBoundary.HardLineBreak)
                {
                    _writer.WriteElementString("w", "br", OpenXmlNamespaces.WordprocessingML, null);
                }
                else if (c == TextBoundary.LineNumber)
                {
                }
                else if (c == TextBoundary.NonBreakingHyphen)
                {
                }
                else if (c == TextBoundary.NonBreakingSpace)
                {
                }
                else if (c == TextBoundary.NonRequiredHyphen)
                {
                }
                else if (c == TextBoundary.PageBreakOrSectionMark)
                {
                }
                else if (c == TextBoundary.ParagraphEnd)
                {
                    //end text
                    _writer.WriteEndElement();
                    //end run
                    _writer.WriteEndElement();
                    //end paragraph
                    _writer.WriteEndElement();

                    startNewParagraph(i, visited);
                }
                else if (c == TextBoundary.Picture)
                {
                }
                else if (c == TextBoundary.Tab)
                {
                    _writer.WriteElementString("w", "tab", OpenXmlNamespaces.WordprocessingML, null);
                }
                else if((int)c < 32)
                {
                    //this is a workaround until fSpec chars are implemented.
                    //characters with special meanings should not be written at all.
                }
                else if (c != '\uFFFF' && !suppressNextChar)
                {
                    _writer.WriteString(new string(c, 1));
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndDocument();
        }

        private void startNewParagraph(int i, WordDocument doc)
        {
            //start new paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //load the matching paragraph properties
            int fc = doc.FIB.fcMin + i;
            if(i!=0)
               fc += 1;
            ParagraphProperties pap = new ParagraphProperties(fc, doc.FIB, doc.WordDocumentStream, doc.TableStream);

            //write the paragraph properties
            _writer.WriteStartElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
            string styleId = StyleSheetMapping.MakeStyleId(doc.Styles.Styles[pap.istd].xstzName);
            _writer.WriteStartElement("w", "pStyle", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, styleId);
            _writer.WriteEndElement();
            pap.Convert(new ParagraphPropertiesMapping(_writer));
            _writer.WriteEndElement();

            //start new run
            _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);
            //start new text
            _writer.WriteStartElement("w", "t", OpenXmlNamespaces.WordprocessingML);
        }
    }
}
