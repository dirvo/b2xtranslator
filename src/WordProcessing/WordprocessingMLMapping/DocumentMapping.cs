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
        private int _papxIndex, _nextPapxFc;
        private WordDocument _doc;
        private List<Int32> _allPapxOffsets;
        private List<ParagraphPropertyExceptions> _allPapx;

        public DocumentMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(WordDocument doc)
        {
            _doc = doc;

            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "document", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "body", OpenXmlNamespaces.WordprocessingML);

            Int32 fcMin = _doc.FIB.fcMin;
            Int32 fcMax = _doc.FIB.fcMin + _doc.FIB.ccpText;

            _doc = doc;
            _allPapx = FormattedDiskPagePAPX.GetParagraphPropertyExceptions(fcMin, fcMax, doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allPapxOffsets = FormattedDiskPagePAPX.GetFileCharacterPositions(fcMin, fcMax, doc.FIB, doc.WordDocumentStream, doc.TableStream);
            _allPapxOffsets.Add(fcMax);

            Int32 cp = 0;
            while (cp < _doc.Text.Count)
            {
                Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
                ParagraphPropertyExceptions papx = _allPapx[_papxIndex];
                TableInfo tai = new TableInfo(papx);
                _nextPapxFc = _allPapxOffsets[_papxIndex];
                
                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    cp = writeTable(cp);
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(papx, cp);
                }

                //If this FC started a new PAPX, take next PAPX next time.
                if (fc == _nextPapxFc && _papxIndex < (_allPapx.Count-1))
                {
                    _papxIndex++;
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndDocument();
        }

        /// <summary>
        /// Writes the table starts at the given cp value
        /// </summary>
        /// <param name="cp">The cp at where the table begins</param>
        /// <returns>The character pointer to the first character after this table</returns>
        private Int32 writeTable(Int32 initialCp)
        {
            //initial values
            Int32 cp = initialCp;
            ParagraphPropertyExceptions papx = _allPapx[_papxIndex];
            TableInfo tai = new TableInfo(papx);

            //start table
            _writer.WriteStartElement("w", "tbl", OpenXmlNamespaces.WordprocessingML);

            //find the first row end PAPX, because it holds the TAP SPRMs
            int tapxIndex = _papxIndex;
            ParagraphPropertyExceptions tablePapx = _allPapx[_papxIndex];
            while (!(new TableInfo(tablePapx).fTtp))
            {
                tablePapx = _allPapx[tapxIndex];
                tapxIndex++;
            }
            //cast it to a TAPX and convert it
            TablePropertyExceptions tapx = new TablePropertyExceptions(tablePapx);
            tapx.Convert(new TablePropertiesMapping(_writer));

            //convert all rows
            while (tai.fInTable)
            {
                cp = writeTableRow(cp);

                //each row has it's own PAPX
                _papxIndex++;
                papx = _allPapx[_papxIndex];
                tai = new TableInfo(papx);
            }

            //close w:tbl
            _writer.WriteEndElement();

            return cp;
        }

        /// <summary>
        /// Writes the table row that starts at the given cp value and ends at the next row end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the row begins</param>
        /// <returns>The character pointer to the first character after this row</returns>
        private Int32 writeTableRow(Int32 initialCp)
        {
            Int32 cp = initialCp;
            ParagraphPropertyExceptions papx = _allPapx[_papxIndex];
            TableInfo  tai = new TableInfo(papx);

            //start w:tr
            _writer.WriteStartElement("w", "tr", OpenXmlNamespaces.WordprocessingML);

            while (!(_doc.Text[cp] == TextBoundary.CellOrRowMark && tai.fTtp))
            {
                cp = writeTableCell(cp);

                //each cell has it's own PAPX
                _papxIndex++;
                papx = _allPapx[_papxIndex];
                tai = new TableInfo(papx);
            }

            //end w:tr
            _writer.WriteEndElement();

            //skip the row end mark
            cp++;

            return cp;
        }

        /// <summary>
        /// Writes the table cell that starts at the given cp value and ends at the next cell end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the cell begins</param>
        /// <returns>The character pointer to the first character after this cell</returns>
        private Int32 writeTableCell(Int32 initialCp)
        {
            Int32 cp = initialCp;

            //start w:tc
            _writer.WriteStartElement("w", "tc", OpenXmlNamespaces.WordprocessingML);

            //find cell end
            Int32 cpCellEnd = initialCp;
            while (_doc.Text[cpCellEnd] != TextBoundary.CellOrRowMark)
            {
                cpCellEnd++;
            }

            //copy the chars of the cell
            List<char> remainingChars = _doc.Text.GetRange(initialCp, cpCellEnd - initialCp);
            while(remainingChars.Contains('\r'))
            {
                Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
                //if this paragraph has a new PAPX
                if (fc == _allPapxOffsets[_papxIndex + 1])
                    _papxIndex++;

                cp = writeParagraph(_allPapx[_papxIndex], cp);
                remainingChars = _doc.Text.GetRange(cp, cpCellEnd - cp);
            }

            //if the remaining chars have an own PAPX
            Int32 fcRemaining = _doc.PieceTable.FileCharacterPositions[cp];
            if (fcRemaining == _allPapxOffsets[_papxIndex + 1])
                _papxIndex++;

            //write the remaining chars as paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);
            _allPapx[_papxIndex].Convert(new ParagraphPropertiesMapping(_writer, _doc.Styles));

            //write all runs for the remaining chars
            Int32 fcCellEnd = _doc.PieceTable.FileCharacterPositions[cpCellEnd];

            //get all CHPX between these boundaries
            List<CharacterPropertyExceptions> chpxs = FormattedDiskPageCHPX.GetCharacterPropertyExceptions(
                fcRemaining,
                fcCellEnd,
                _doc.FIB,
                _doc.WordDocumentStream,
                _doc.TableStream);
            List<Int32> chpxFcs = FormattedDiskPageCHPX.GetFileCharacterPositions(
                fcRemaining,
                fcCellEnd,
                _doc.FIB,
                _doc.WordDocumentStream,
                _doc.TableStream);
            chpxFcs.Add(fcCellEnd);

            //write runs for all CHPX
            for (int i = 0; i < chpxs.Count; i++)
            {
                //get the chars of this CHPX
                int fcChpxStart = chpxFcs[i];
                int fcChpxEnd = chpxFcs[i + 1];

                //it's the first chpx and it starts before the paragraph
                if (i == 0 && fcChpxStart < fcRemaining)
                {
                    //so use the FC of the paragraph
                    fcChpxStart = fcRemaining;
                }

                //it's the last chpx and it exceeds the paragraph
                if (i == (chpxs.Count - 1) && fcChpxEnd > fcCellEnd)
                {
                    //so use the FC of the paragraph
                    fcChpxEnd = fcCellEnd;
                }

                //read the chars that are formatted via this CHPX
                List<char> chpxChars = _doc.PieceTable.GetChars(fcChpxStart, fcChpxEnd, _doc.WordDocumentStream);
                
                //write the run
                if (chpxChars.Count > 0)
                    writeRun(chpxChars, chpxs[i]);

                //increase the pointer
                cp += chpxChars.Count;
            }

            //end w:p
            _writer.WriteEndElement();

            //end w:tc
            _writer.WriteEndElement();

            //skip cell end mark
            cp++;

            return cp;
        }


        /// <summary>
        /// Writes a Paragraph for the given PAPX
        /// </summary>
        private Int32 writeParagraph(ParagraphPropertyExceptions papx, Int32 cp) 
        {
            //start paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //write properties
            papx.Convert(new ParagraphPropertiesMapping(_writer, _doc.Styles));

            //search the paragraph end
            Int32 cpParaEnd = cp;
            while (_doc.Text[cpParaEnd] != TextBoundary.ParagraphEnd)
            {
                cpParaEnd++;
            }
            cpParaEnd++;

            //get the physical boundaries (FC) of that paragraph
            Int32 fcStart = _doc.PieceTable.FileCharacterPositions[cp];
            Int32 fcEnd = _doc.PieceTable.FileCharacterPositions[cpParaEnd];

            //get all CHPX between these boundaries
            List<CharacterPropertyExceptions> chpxs = FormattedDiskPageCHPX.GetCharacterPropertyExceptions(
                fcStart,
                fcEnd,
                _doc.FIB,
                _doc.WordDocumentStream,
                _doc.TableStream);
            List<Int32> chpxFcs = FormattedDiskPageCHPX.GetFileCharacterPositions(
                fcStart,
                fcEnd,
                _doc.FIB,
                _doc.WordDocumentStream,
                _doc.TableStream);
            chpxFcs.Add(fcEnd);

            //write runs for all CHPX
            for (int i = 0; i < chpxs.Count; i++)
            {
                //get the chars of this CHPX
                int fcChpxStart = chpxFcs[i];
                int fcChpxEnd = chpxFcs[i + 1];

                //it's the first chpx and it starts before the paragraph
                if (i == 0 && fcChpxStart < fcStart)
                {
                    //so use the FC of the paragraph
                    fcChpxStart = fcStart;
                }
                
                //it's the last chpx and it exceeds the paragraph
                if (i == (chpxs.Count - 1) && fcChpxEnd > fcEnd)
                {
                    //so use the FC of the paragraph
                    fcChpxEnd = fcEnd;
                }

                //read the chars that are formatted via this CHPX
                List<char> chpxChars = _doc.PieceTable.GetChars(fcChpxStart, fcChpxEnd, _doc.WordDocumentStream);

                //write the run
                if(chpxChars.Count > 0)
                    writeRun(chpxChars, chpxs[i]);
            }

            //end paragraph
            _writer.WriteEndElement();

            return cpParaEnd++;
        }

        /// <summary>
        /// Writes a run with the given characters and CHPX
        /// </summary>
        private void writeRun(List<char> chars, CharacterPropertyExceptions chpx)
        {
            //start run
            _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);

            //write properties
            chpx.Convert(new CharacterPropertiesMapping(_writer, _doc.Styles, _doc.FontTable));

            //start text
            _writer.WriteStartElement("w", "t", OpenXmlNamespaces.WordprocessingML);

            if ((int)chars[0] == 32 || (int)chars[chars.Count - 1] == 32)
            {
                _writer.WriteAttributeString("xml", "space", "", "preserve");
            }

            //write text
            foreach(char c in chars)
            {
                if (c == TextBoundary.Tab)
                {
                    _writer.WriteElementString("w", "tab", OpenXmlNamespaces.WordprocessingML, "");
                }
                else if (c == TextBoundary.HardLineBreak)
                {
                    _writer.WriteElementString("w", "br", OpenXmlNamespaces.WordprocessingML, "");
                }
                else if(c == TextBoundary.Picture)
                {
                    _writer.WriteString("[PICTURE]");
                }
                else if (c == TextBoundary.ParagraphEnd)
                {
                    //do nothing
                }
                else if (c == TextBoundary.PageBreakOrSectionMark)
                {
                    //do nothing
                }
                else if (c == TextBoundary.ColumnBreak)
                {
                    //do nothing
                }
                else if (c == TextBoundary.FieldBeginMark)
                {
                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "begin");
                    _writer.WriteEndElement();
                }
                else if (c == TextBoundary.FieldSeperator)
                {
                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "separate");
                    _writer.WriteEndElement();
                }
                else if (c == TextBoundary.FieldEndMark)
                {
                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "end");
                    _writer.WriteEndElement();
                }
                else if  ((int)c > 31 && (int)c != 0xFFFF)
                {
                     _writer.WriteString(new String(c, 1));
                }
            }

            //end run
            _writer.WriteEndElement();

            //end run
            _writer.WriteEndElement();
        }
    }
}
