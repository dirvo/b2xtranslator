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
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class DocumentMapping : 
        AbstractOpenXmlMapping,
        IMapping<WordDocument>
    {
        private WordDocument _doc;
        private MainDocumentPart _docPart;
        private XmlWriterSettings _xws;

        /// <summary>
        /// A dictionary that contains all PAPX of the document.<br/>
        /// The key is the FC at which the paragraph starts.<br/>
        /// The value is the PAPX that formats the paragraph.
        /// </summary>
        private Dictionary<Int32, ParagraphPropertyExceptions> _allPapx;

        /// <summary>
        /// A dictionary that contains all SEPX of the document.<br/>
        /// The key is the CP at which sections ends.<br/>
        /// The value is the SEPX that formats the section.
        /// </summary>
        private Dictionary<Int32, SectionPropertyExceptions> _allSepx;

        public DocumentMapping(MainDocumentPart docPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(docPart.GetStream(), xws))
        {
            _xws = xws;
            _docPart = docPart;
        }

        public void Apply(WordDocument doc)
        {
            _doc = doc;

            //build a dictionary of all PAPX
            _allPapx = new Dictionary<Int32, ParagraphPropertyExceptions>();
            for (int i = 0; i < _doc.AllPapxFkps.Count; i++)
            {
                for (int j = 0; j < _doc.AllPapxFkps[i].grppapx.Length; j++)
                {
                    _allPapx.Add(_doc.AllPapxFkps[i].rgfc[j], _doc.AllPapxFkps[i].grppapx[j]);
                }
            }

            //build a dictionary of all SEPX
            _allSepx = new Dictionary<Int32, SectionPropertyExceptions>();
            for (int i = 0; i < _doc.SectionTable.grpsepx.Length; i++)
            {
                _allSepx.Add(_doc.SectionTable.rgfc[i+1], _doc.SectionTable.grpsepx[i]);
            }

            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "document", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "body", OpenXmlNamespaces.WordprocessingML);

            Int32 cp = 0;
            while (cp < _doc.Text.Count)
            {
                Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
                ParagraphPropertyExceptions papx = findValidPapx(fc);
                TableInfo tai = new TableInfo(papx);

                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    cp = writeTable(cp);
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
                }
            }

            //write the section properties of the body with the last SEPX
            SectionPropertyExceptions lastSepx = _doc.SectionTable.grpsepx[_doc.SectionTable.grpsepx.Length - 1];
            lastSepx.Convert(new SectionPropertiesMapping(_writer));

            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        /// <summary>
        /// Writes the table starts at the given cp value
        /// </summary>
        /// <param name="cp">The cp at where the table begins</param>
        /// <returns>The character pointer to the first character after this table</returns>
        private Int32 writeTable(Int32 initialCp)
        {
            Int32 cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo tai = new TableInfo(papx);

            //start table
            _writer.WriteStartElement("w", "tbl", OpenXmlNamespaces.WordprocessingML);

            //find first row end TAPX
            TablePropertyExceptions row1Tapx = findRowEndTapx(cp);

            //Convert it
            row1Tapx.Convert(new TablePropertiesMapping(_writer, _doc.Styles));

            //convert all rows
            while (tai.fInTable)
            {
                cp = writeTableRow(cp);

                fc = _doc.PieceTable.FileCharacterPositions[cp];
                papx = findValidPapx(fc);
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
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo  tai = new TableInfo(papx);

            //start w:tr
            _writer.WriteStartElement("w", "tr", OpenXmlNamespaces.WordprocessingML);

            //convert the properties
            TablePropertyExceptions tapx = findRowEndTapx(cp);
            tapx.Convert(new TableRowPropertiesMapping(_writer));

            int cellIndex = 0;
            while (!(_doc.Text[cp] == TextBoundary.CellOrRowMark && tai.fTtp) && tai.fInTable)
            {
                cp = writeTableCell(cp, cellIndex, tapx);
                cellIndex++;

                //each cell has it's own PAPX
                fc = _doc.PieceTable.FileCharacterPositions[cp];
                papx = findValidPapx(fc);
                tai = new TableInfo(papx);
            }

            //end w:tr
            _writer.WriteEndElement();

            //skip the row end mark
            cp++;

            return cp;
        }

        /// <summary>
        /// Finds the TAPX that formats the next row end mark.
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        private TablePropertyExceptions findRowEndTapx(int initialCp)
        {
            int cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo tai = new TableInfo(papx);

            while (tai.fTtp==false && tai.fInTable==true)
            {
                while (_doc.Text[cp] != TextBoundary.CellOrRowMark)
                {
                    cp++;
                }
                fc = _doc.PieceTable.FileCharacterPositions[cp];
                papx = findValidPapx(fc);
                tai = new TableInfo(papx);
                cp++;
            }

            return new TablePropertyExceptions(papx, _doc.DataStream);
        }

        /// <summary>
        /// Writes the table cell that starts at the given cp value and ends at the next cell end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the cell begins</param>
        /// <param name="cellIndex">The index of this cell. The first cell's index should be 0</param>
        /// <param name="initialCp">The TAPX that formats the row to which the cell belongs</param>
        /// <returns>The character pointer to the first character after this cell</returns>
        private Int32 writeTableCell(Int32 initialCp, int cellIndex, TablePropertyExceptions tapx)
        {
            Int32 cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];

            //start w:tc
            _writer.WriteStartElement("w", "tc", OpenXmlNamespaces.WordprocessingML);

            //find cell end
            Int32 cpCellEnd = initialCp;
            while (_doc.Text[cpCellEnd] != TextBoundary.CellOrRowMark)
            {
                cpCellEnd++;
            }

            //convert the properties
            tapx.Convert(new TableCellPropertiesMapping(_writer, cellIndex));

            //copy the chars of the cell
            List<char> remainingChars = _doc.Text.GetRange(initialCp, cpCellEnd - initialCp);

            //write paragaphs for each paragraph in the cell
            while(remainingChars.Contains('\r'))
            {
                fc = _doc.PieceTable.FileCharacterPositions[cp];
                cp = writeParagraph(cp);
                remainingChars = _doc.Text.GetRange(cp, cpCellEnd - cp);
            }

            //write the remaining chars as own paragraph
            cp = writeParagraph(cp, cpCellEnd, false);

            //end w:tc
            _writer.WriteEndElement();

            //skip cell end mark
            cp++;

            return cp;
        }

        /// <summary>
        /// Writes a Paragraph that starts at the given cp and 
        /// ends at the next paragraph end mark or section mark
        /// </summary>
        /// <param name="cp"></param>
        private Int32 writeParagraph(Int32 cp) 
        {
            //search the paragraph end
            Int32 cpParaEnd = cp;
            while (_doc.Text[cpParaEnd] != TextBoundary.ParagraphEnd && 
                _doc.Text[cpParaEnd] != TextBoundary.PageBreakOrSectionMark)
            {
                cpParaEnd++;
            }


            //end was a section break
            if (_doc.Text[cpParaEnd] == TextBoundary.PageBreakOrSectionMark)
            {
                cpParaEnd++;
                return writeParagraph(cp, cpParaEnd, true);
            }
            else
            {
                cpParaEnd++;
                return writeParagraph(cp, cpParaEnd, false);
            }
        }

        /// <summary>
        /// Writes a Paragraph that starts at the given cpStart and 
        /// ends at the given cpEnd
        /// </summary>
        /// <param name="cpStart"></param>
        /// <param name="cpEnd"></param>
        /// <param name="sectionEnd">Set if this paragraph is the last paragraph of a section</param>
        /// <returns></returns>
        private Int32 writeParagraph(Int32 initialCp, Int32 cpEnd, bool sectionEnd)
        {
            Int32 cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            Int32 fcEnd = _doc.PieceTable.FileCharacterPositions[cpEnd];
            ParagraphPropertyExceptions papx = findValidPapx(fc);

            //start paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //check for section properties
            if (sectionEnd)
            {
                //this is the last paragraph of this section
                //write properties with section properties
                papx.Convert(new ParagraphPropertiesMapping(_writer, _doc.Styles, findValidSepx(cpEnd)));
            }
            else
            {
                //write properties
                papx.Convert(new ParagraphPropertiesMapping(_writer, _doc.Styles));
            }

            //get all CHPX between these boundaries to determine the count of runs
            List<CharacterPropertyExceptions> chpxs = FormattedDiskPageCHPX.GetCharacterPropertyExceptions(
                fc,
                fcEnd,
                _doc.FIB,
                _doc.WordDocumentStream,
                _doc.TableStream);
            List<Int32> chpxFcs = FormattedDiskPageCHPX.GetFileCharacterPositions(
                fc,
                fcEnd,
                _doc.FIB,
                _doc.WordDocumentStream,
                _doc.TableStream);
            chpxFcs.Add(fcEnd);

            //write a rus for each CHPX
            for (int i = 0; i < chpxs.Count; i++)
            {
                //get the chars of this CHPX
                int fcChpxStart = chpxFcs[i];
                int fcChpxEnd = chpxFcs[i + 1];

                //it's the first chpx and it starts before the paragraph
                if (i == 0 && fcChpxStart < fc)
                {
                    //so use the FC of the paragraph
                    fcChpxStart = fc;
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
                if (chpxChars.Count > 0)
                {
                    cp = writeRun(chpxChars, chpxs[i], cp);
                }
            }

            //end paragraph
            _writer.WriteEndElement();

            return cpEnd++;
        }

        /// <summary>
        /// Writes a run with the given characters and CHPX
        /// </summary>
        private Int32 writeRun(List<char> chars, CharacterPropertyExceptions chpx, Int32 initialCp)
        {
            Int32 cp = initialCp;

            //start run
            _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);

            //write properties
            chpx.Convert(new CharacterPropertiesMapping(_writer, _doc.Styles, _doc.FontTable));

            if (chars.Count == 1 && chars[0] == TextBoundary.Picture)
            {
                //its a picture
                PictureDescriptor pict = new PictureDescriptor(chpx, _doc.DataStream);

                //sometimes there is a picture mark without a picture,
                //do not convert these marks
                if (pict.mfp.mm > 98)
                {
                    ImagePart imgPart = copyPicture(pict);
                    pict.Convert(new PictureMapping(_writer, imgPart));
                }

                cp++;
            }
            else
            {
                cp = writeText(chars, cp);
            }

            //end run
            _writer.WriteEndElement();

            return cp;
        }

        /// <summary>
        /// Writes the given text to the document
        /// </summary>
        /// <param name="chars"></param>
        private Int32 writeText(List<char> chars, Int32 initialCp)
        {
            Int32 cp = initialCp;

            //start text
            _writer.WriteStartElement("w", "t", OpenXmlNamespaces.WordprocessingML);

            if ((int)chars[0] == 32 || (int)chars[chars.Count - 1] == 32)
            {
                _writer.WriteAttributeString("xml", "space", "", "preserve");
            }

            //write text
            foreach (char c in chars)
            {
                if (c == TextBoundary.Tab)
                {
                    _writer.WriteElementString("w", "tab", OpenXmlNamespaces.WordprocessingML, "");
                }
                else if (c == TextBoundary.HardLineBreak)
                {
                    _writer.WriteElementString("w", "br", OpenXmlNamespaces.WordprocessingML, "");
                }
                else if (c == TextBoundary.Picture)
                {
                    //do nothing
                    //see picture conversion above
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
                    _writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "column");
                    _writer.WriteEndElement();
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
                else if ((int)c > 31 && (int)c != 0xFFFF)
                {
                    _writer.WriteString(new String(c, 1));
                }

                cp++;
            }

            //end text
            _writer.WriteEndElement();

            return cp;
        }

        /// <summary>
        /// Copies the picture from the binary stream to the zip archive 
        /// and creates the relationships for the image.
        /// </summary>
        /// <param name="pict">The PictureDescriptor</param>
        /// <returns>The created ImagePart</returns>
        private ImagePart copyPicture(PictureDescriptor pict)
        {
            //create the image part
            ImagePart imgPart = null;
            switch (pict.Type)
            {
                case PictureDescriptor.PictureType.jpg:
                    imgPart = _docPart.AddImagePart(ImagePartType.Jpeg);
                    break;
                case PictureDescriptor.PictureType.png:
                    imgPart = _docPart.AddImagePart(ImagePartType.Png);
                    break;
                case PictureDescriptor.PictureType.wmf:
                    imgPart = _docPart.AddImagePart(ImagePartType.Wmf);
                    break;
                default:
                    imgPart = _docPart.AddImagePart(ImagePartType.Png);
                    break;
            }

            //write the picture
            imgPart.GetStream().Write(pict.Picture, 0, pict.Picture.Length);

            return imgPart;
        }

        /// <summary>
        /// Checks if the CHPX is special
        /// </summary>
        /// <param name="chpx">The CHPX</param>
        /// <returns></returns>
        private bool isSpecial(CharacterPropertyExceptions chpx)
        {
            bool ret = false;
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                if (sprm.OpCode == 0x6A03 || sprm.OpCode == 0x6A12)
                {
                    //special picture
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == 0x6A09)
                {
                    //special symbol
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == 0x0855)
                {
                    //special value
                    ret = DIaLOGIKa.b2xtranslator.DocFileFormat.Utils.ByteToBool(sprm.Arguments[0]);
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Finds the PAPX that is valid for the given FC.
        /// </summary>
        /// <param name="fc"></param>
        /// <returns></returns>
        private ParagraphPropertyExceptions findValidPapx(Int32 fc)
        {
            ParagraphPropertyExceptions ret = null;

            while (ret == null)
            {
                try
                {
                    ret = _allPapx[fc];
                }
                catch (KeyNotFoundException)
                {
                    fc--;
                }
            }

            return ret;
        }

        /// <summary>
        /// Finds the SEPX that is valid for the given CP.
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        private SectionPropertyExceptions findValidSepx(Int32 cp)
        {
            SectionPropertyExceptions ret = null;

            try
            {
                ret = _allSepx[cp];
            }
            catch (KeyNotFoundException)
            {
                //there is no Sepx at this position, 
                //so the previous sepx is valid for this cp

                Int32 lastKey = _doc.SectionTable.rgfc[1];
                foreach (Int32 key in _allSepx.Keys)
                {
                    if (cp > lastKey && cp < key)
                    {
                        ret = _allSepx[lastKey];
                        break;
                    }
                }
            }

            return ret;
        }
    }
}
