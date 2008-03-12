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
    public class DocumentMapping : 
        AbstractOpenXmlMapping,
        IMapping<WordDocument>
    {

        private List<FormattedDiskPagePAPX> _papxFkps;
        private List<FormattedDiskPageCHPX> _chpxFkps;
        private List<Int32> _allPapxOffsets, _allChpxOffsets;
        private List<ParagraphPropertyExceptions> _allPapx;
        private List<CharacterPropertyExceptions> _allChpx;

        private bool _charIsSpecial;
        private WordDocument _doc;

        private int _chpxIndex = 0;
        private int _papxIndex = 0;

        public DocumentMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(WordDocument doc)
        {
            _doc = doc;

            int fcMin = doc.FIB.fcMin;
            int fcMax = doc.FIB.fcMin + doc.FIB.ccpText;

            _allChpxOffsets = FormattedDiskPageCHPX.GetFileCharacterPositions(fcMin, fcMax, doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allChpx = FormattedDiskPageCHPX.GetCharacterPropertyExceptions(fcMin, fcMax, doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allPapxOffsets = FormattedDiskPagePAPX.GetFileCharacterPositions(fcMin, fcMax, doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allPapx = FormattedDiskPagePAPX.GetParagraphPropertyExceptions(fcMin, fcMax, doc.FIB, doc.WordDocumentStream, doc.TableStream);

            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "document", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "body", OpenXmlNamespaces.WordprocessingML);

            bool suppressNextChar = false;

            int fc = doc.FIB.fcMin;
            int nextChpxOffset = fc;
            int nextPapxOffset = fc;

            for (int i = 0; i < doc.Text.Count; i++)
            {
                //PAPX starts here
                if (fc == nextPapxOffset)
                {
                    startNewParagraph();

                    //but no chpx starts here:
                    if (nextPapxOffset != nextChpxOffset)
                    {
                        //That means, that the text is formatted with the same 
                        //chpx as the previous paragraph.
                        //write new run, but don't increase the index
                        startNewRun();
                    }

                    //increase index for the next time
                    _papxIndex++;
                    nextPapxOffset = _allPapxOffsets[_papxIndex];
                }

                //CHPX starts here
                if (fc == nextChpxOffset)
                {
                    startNewRun();

                    //increase index for the next time
                    _chpxIndex++;
                    nextChpxOffset = _allChpxOffsets[_chpxIndex];
                }

                #region char handling
                char c = doc.Text[i];

                //check the char
                if (c == TextBoundary.BreakingHyphen)
                {
                }
                else if (c == TextBoundary.CellOrRowMark)
                {
                }
                else if (c == TextBoundary.ColumnBreak)
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
                else if (c == TextBoundary.HardLineBreak)
                {
                    _writer.WriteElementString("w", "br", OpenXmlNamespaces.WordprocessingML, null);
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
                    //see pararagraph end handling below
                }
                else if (c == TextBoundary.Tab)
                {
                    _writer.WriteElementString("w", "tab", OpenXmlNamespaces.WordprocessingML, null);
                }
                else if (_charIsSpecial && c == TextBoundary.CurrentPageNumber)
                {
                }
                else if (_charIsSpecial && c == TextBoundary.Picture)
                {
                }
                else if (_charIsSpecial && c == TextBoundary.AutoNumberedFootnoteReference)
                {
                }
                else if (_charIsSpecial && c == TextBoundary.FootnoteContinuation)
                {
                }
                else if (_charIsSpecial && c == TextBoundary.FootnoteSeparator)
                {
                }
                else if (_charIsSpecial && c == TextBoundary.AnnotationReference)
                {
                }
                else if (_charIsSpecial && c == TextBoundary.LineNumber)
                {
                }
                else if (_charIsSpecial && c == TextBoundary.HandAnnotationPicture)
                {
                }
                else if (c != '\uFFFF' && !suppressNextChar && (int)c > 31)
                {
                    _writer.WriteString(new string(c, 1));
                }
                #endregion

                //CHPX ends here
                if (fc == nextChpxOffset-1)
                {
                    //close w:t
                    _writer.WriteEndElement();
                    //close w:r
                    _writer.WriteEndElement();
                }

                //PAPX ends here
                if(fc == nextPapxOffset-1)
                {
                    //but no CHPX ended here
                    if (nextPapxOffset != nextChpxOffset)
                    {
                        //close w:t
                        _writer.WriteEndElement();
                        //close w:r
                        _writer.WriteEndElement();
                    }

                    //close w:p
                    _writer.WriteEndElement();
                }

                fc++;
            }         

            //end body
            _writer.WriteEndElement();
            //end document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();
        }

        private void startNewParagraph()
        {
            //write new paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //get papx
            ParagraphPropertyExceptions papx = _allPapx[_papxIndex];

            //convert it
            papx.Convert(new ParagraphPropertiesMapping(_writer, _doc.Styles));
        }

        private void startNewRun()
        {
            //write new run
            _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);
            //write new text
            _writer.WriteStartElement("w", "t", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("xml", "space", "", "preserve");

            //get chpx
            CharacterPropertyExceptions chpx = _allChpx[_chpxIndex];

            //check if the fSpec flag is set
            _charIsSpecial = false;
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                if (sprm.OpCode == 0x0855 && sprm.Arguments[0] == 1)
                    _charIsSpecial = true;
            }

            //convert the chpx
            chpx.Convert(new CharacterPropertiesMapping(_writer, _doc.Styles, _doc.FontTable));
        }
    }
}
