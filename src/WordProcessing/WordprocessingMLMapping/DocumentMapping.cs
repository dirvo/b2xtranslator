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

            _allChpxOffsets = FormattedDiskPageCHPX.GetAllFCs(doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allChpx = FormattedDiskPageCHPX.GetAllCHPX(doc.FIB.fcMin, doc.FIB.fcMin+doc.FIB.ccpText, doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allPapxOffsets = FormattedDiskPagePAPX.GetAllFCs(doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allPapx = FormattedDiskPagePAPX.GetAllPAPX(doc.FIB.fcMin, doc.FIB.fcMin + doc.FIB.ccpText, doc.FIB, doc.WordDocumentStream, doc.TableStream);

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
                else if((int)c < 32)
                {
                    //this is a workaround until fSpec chars are implemented.
                    //characters with special meanings should not be written at all.
                }
                else if (c != '\uFFFF' && !suppressNextChar)
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

        private static void drawTextProgressBar(int progress, int total)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 32;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 30.0f / total;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw unfilled part
            for (int i = position; i < 31; i++)
            {
                Console.BackgroundColor = ConsoleColor.Black;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            Console.CursorLeft = 35;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write(progress.ToString() + " of " + total.ToString() + "    "); //blanks at the end remove any excess
        }

        private void startNewParagraph()
        {
            //write new paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //get papx
            ParagraphPropertyExceptions papx = _allPapx[_papxIndex];

            //convert it
            papx.Convert(new ParagraphPropertiesMapping(_writer, _doc.Styles));

            #region oldcode
            ////find the matching PAPX in the FKPs
            //foreach (FormattedDiskPagePAPX fkp in _papxFkps)
            //{
            //   for (int i = 0; i < fkp.grppapx.Length; i++)
            //   {
            //       if (fkp.rgfc[i] == fc)
            //       {
            //           ParagraphPropertyExceptions papx = fkp.grppapx[i];

            //           //convert the properties
            //           papx.Convert(new ParagraphPropertiesMapping(_writer, doc.Styles));

            //           //set offset of next paragraph
            //           _nextParaOffset = fkp.rgfc[i + 1];

            //           break;
            //       }
            //   }
            //}
            #endregion
        }

        private void startNewRun()
        {
            //drawTextProgressBar(_chpxIndex + 1, _allChpx.Count);

            //write new run
            _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);
            //write new text
            _writer.WriteStartElement("w", "t", OpenXmlNamespaces.WordprocessingML);

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
            chpx.Convert(new CharacterPropertiesMapping(_writer, _doc.Styles));

            #region oldcode
            ////find the matching CHPX in the FKPs
            //foreach (FormattedDiskPageCHPX fkp in _chpxFkps)
            //{
            //    for (int i = 0; i < fkp.grpchpx.Length; i++)
            //    {
            //        if (fkp.rgfc[i] == fc)
            //        {
            //            CharacterPropertyExceptions chpx = fkp.grpchpx[i];

            //            //check if the fSpec flag is set
            //            _charIsSpecial = false;
            //            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            //            {
            //                if (sprm.OpCode == 0x0855 && sprm.Arguments[0] == 1)
            //                    _charIsSpecial = true;
            //            }
                        
            //            //convert the properties
            //            chpx.Convert(new CharacterPropertiesMapping(_writer, doc.Styles));

            //            //set offset of next run
            //            _nextRunOffset = fkp.rgfc[i + 1];

            //            break;
            //        }
            //    }
            //}
            #endregion
        }
    }
}
