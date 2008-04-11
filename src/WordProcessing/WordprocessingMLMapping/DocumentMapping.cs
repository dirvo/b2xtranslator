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
using DIaLOGIKa.b2xtranslator.Utils;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class DocumentMapping : 
        AbstractOpenXmlMapping,
        IMapping<WordDocument>
    {
        private WordDocument _doc;
        private MainDocumentPart _docPart;
        private XmlWriterSettings _xws;
        private ParagraphPropertyExceptions _lastValidPapx;
        private SectionPropertyExceptions _lastValidSepx;

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

            //build a dictionaries of all PAPX
            _allPapx = new Dictionary<Int32, ParagraphPropertyExceptions>();
            for (int i = 0; i < _doc.AllPapxFkps.Count; i++)
            {
                for (int j = 0; j < _doc.AllPapxFkps[i].grppapx.Length; j++)
                {
                     _allPapx.Add(_doc.AllPapxFkps[i].rgfc[j], _doc.AllPapxFkps[i].grppapx[j]);
                }
            }
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];

            //build a dictionaries of all SEPX
            _allSepx = new Dictionary<Int32, SectionPropertyExceptions>();
            for (int i = 0; i < _doc.SectionTable.grpsepx.Length; i++)
            {
                _allSepx.Add(_doc.SectionTable.rgfc[i + 1], _doc.SectionTable.grpsepx[i]);
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

        #region TableConversion

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

            //build the table grid
            List<Int16> grid = buildTableGrid(cp);

            //find first row end
            Int32 fcRowEnd = findRowEndFc(cp);
            TablePropertyExceptions row1Tapx = new TablePropertyExceptions(findValidPapx(fcRowEnd), _doc.DataStream);

            //start table
            _writer.WriteStartElement("w", "tbl", OpenXmlNamespaces.WordprocessingML);

            //Convert it
            row1Tapx.Convert(new TablePropertiesMapping(_writer, _doc.Styles, grid));

            //convert all rows
            while (tai.fInTable)
            {
                cp = writeTableRow(cp, grid);

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
        private Int32 writeTableRow(Int32 initialCp, List<Int16> grid)
        {
            Int32 cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo  tai = new TableInfo(papx);

            //start w:tr
            _writer.WriteStartElement("w", "tr", OpenXmlNamespaces.WordprocessingML);

            //convert the properties
            Int32 fcRowEnd = findRowEndFc(cp);
            TablePropertyExceptions tapx = new TablePropertyExceptions(findValidPapx(fcRowEnd), _doc.DataStream);
            List<CharacterPropertyExceptions> chpxs = _doc.GetCharacterPropertyExceptions(fcRowEnd, fcRowEnd + 1);

            tapx.Convert(new TableRowPropertiesMapping(_writer, chpxs[0]));

            int gridIndex = 0;
            int cellIndex = 0;
            while (!(_doc.Text[cp] == TextBoundary.CellOrRowMark && tai.fTtp) && tai.fInTable)
            {
                cp = writeTableCell(cp, tapx, grid, ref gridIndex, cellIndex);
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
        /// Builds a list that contains the width of the several columns of the table.
        /// </summary>
        /// <param name="initialCp"></param>
        /// <returns></returns>
        private List<Int16> buildTableGrid(int initialCp)
        {
            ParagraphPropertyExceptions backup = _lastValidPapx;

            List<Int16> boundaries = new List<Int16>();
            List<Int16> grid = new List<Int16>();
            Int32 cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo tai = new TableInfo(papx);

            Int32 fcRowEnd = findRowEndFc(cp, out cp);
            int maxColumnCount = 0;

            while (tai.fInTable)
            {
                //check all SPRMs of this TAPX
                foreach (SinglePropertyModifier sprm in papx.grpprl)
                {
                    //find the tDef SPRM
                    if(sprm.OpCode == 0xd608)
                    {
                        byte itcMac = sprm.Arguments[0];
                        for (int i = 0; i < itcMac; i++)
                        {
                            Int16 boundary1 = System.BitConverter.ToInt16(sprm.Arguments, 1 + (i * 2));
                            if (!boundaries.Contains(boundary1))
                                boundaries.Add(boundary1);

                            Int16 boundary2 = System.BitConverter.ToInt16(sprm.Arguments, 1 + ((i + 1) * 2));
                            if (!boundaries.Contains(boundary2))
                                boundaries.Add(boundary2);
                        }
                    }
                }

                //get the next papx
                papx = findValidPapx(fcRowEnd);
                tai = new TableInfo(papx);
                fcRowEnd = findRowEndFc(cp, out cp);
            }

            //build the grid based on the boundaries
            boundaries.Sort();
            for (int i = 0; i < boundaries.Count -1; i++)
            {
                grid.Add( (Int16)(boundaries[i+1] - boundaries[i]) );
            }

            _lastValidPapx = backup;
            return grid;
        }

        /// <summary>
        /// Finds the FC of the next row end mark.
        /// </summary>
        /// <param name="initialCp">Some CP before the row end</param>
        /// <param name="rowEndCp">The CP of the next row end mark</param>
        /// <returns>The FC of the next row end mark</returns>
        private Int32 findRowEndFc(int initialCp, out int rowEndCp)
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

            rowEndCp = cp;
            return fc;
        }

        /// <summary>
        /// Finds the FC of the next row end mark.
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        private Int32 findRowEndFc(int initialCp)
        {
            int cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo tai = new TableInfo(papx);

            while (tai.fTtp == false && tai.fInTable == true)
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

            return fc;
        }

        /// <summary>
        /// Writes the table cell that starts at the given cp value and ends at the next cell end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the cell begins</param>
        /// <param name="tapx">The TAPX that formats the row to which the cell belongs</param>
        /// <param name="gridIndex">The index of this cell in the grid</param>
        /// <param name="gridIndex">The grid</param>
        /// <returns>The character pointer to the first character after this cell</returns>
        private Int32 writeTableCell(Int32 initialCp, TablePropertyExceptions tapx, List<Int16> grid, ref int gridIndex, int cellIndex)
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
            cpCellEnd++;

            //convert the properties
            TableCellPropertiesMapping mapping = new TableCellPropertiesMapping(_writer, grid, gridIndex, cellIndex);
            tapx.Convert(mapping);
            gridIndex = gridIndex + mapping.GridSpan;

            //write the paragraphs of the cell
            while (cp < cpCellEnd)
            {
                cp = writeParagraph(cp);
            }

            //end w:tc
            _writer.WriteEndElement();

            return cp;
        }

        #endregion

        #region ParagraphRunConversion

        /// <summary>
        /// Writes a Paragraph that starts at the given cp and 
        /// ends at the next paragraph end mark or section end mark
        /// </summary>
        /// <param name="cp"></param>
        private Int32 writeParagraph(Int32 cp) 
        {
            //search the paragraph end
            Int32 cpParaEnd = cp;
            while (_doc.Text[cpParaEnd] != TextBoundary.ParagraphEnd && 
                _doc.Text[cpParaEnd] != TextBoundary.CellOrRowMark &&
                !(_doc.Text[cpParaEnd] == TextBoundary.PageBreakOrSectionMark && isSectionEnd(cpParaEnd)))
            {
                cpParaEnd++;
            }

            if (_doc.Text[cpParaEnd] == TextBoundary.PageBreakOrSectionMark)
            {
                //there is a page break OR section mark,
                //write the section only if it's a section mark
                bool sectionEnd = isSectionEnd(cpParaEnd);
                cpParaEnd++;
                return writeParagraph(cp, cpParaEnd, sectionEnd);
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

            //get all CHPX between these boundaries to determine the count of runs
            List<CharacterPropertyExceptions> chpxs = _doc.GetCharacterPropertyExceptions(fc, fcEnd);
            List<Int32> chpxFcs = _doc.GetFileCharacterPositions(fc, fcEnd);
            chpxFcs.Add(fcEnd);

            //the last of these CHPX formast the paragraph end mark
            CharacterPropertyExceptions paraEndChpx = chpxs[chpxs.Count-1];

            //start paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //check for section properties
            if (sectionEnd)
            {
                //this is the last paragraph of this section
                //write properties with section properties
                papx.Convert(new ParagraphPropertiesMapping(_writer, _doc, paraEndChpx, findValidSepx(cpEnd)));
            }
            else
            {
                //write properties
                papx.Convert(new ParagraphPropertiesMapping(_writer, _doc, paraEndChpx));
            }

            //write a run for each CHPX
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

            //If it's a deleted run
            if (chpx.IsDeleted)
            {
                _writer.WriteStartElement("w", "del", OpenXmlNamespaces.WordprocessingML);
            }

            //start run
            _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);

            //append rsid
            //_writer.WriteAttributeString("w", "rsidRPr", OpenXmlNamespaces.WordprocessingML, getChpxRsid(chpx).ToString());

            //write properties
            chpx.Convert(new CharacterPropertiesMapping(_writer, _doc));

            if (chars.Count == 1 && chars[0] == TextBoundary.Picture)
            {
                ////its a picture
                //PictureDescriptor pict = new PictureDescriptor(chpx, _doc.DataStream);

                ////sometimes there is a picture mark without a picture,
                ////do not convert these marks (occurs in hyperlinks e.g.)
                //if (pict.mfp.mm > 98)
                //{
                //    ImagePart imgPart = copyPicture(pict);
                //    pict.Convert(new PictureMapping(_writer, imgPart));
                //}

                cp++;
            }
            else
            {
                cp = writeText(chars, cp, chpx.IsDeleted);
            }

            //end run
            _writer.WriteEndElement();

            //If it's a deleted run
            if (chpx.IsDeleted)
            {
                _writer.WriteEndElement();
            }

            return cp;
        }

        /// <summary>
        /// Writes the given text to the document
        /// </summary>
        /// <param name="chars"></param>
        private Int32 writeText(List<char> chars, Int32 initialCp, bool writeDeletedText)
        {
            Int32 cp = initialCp;

            //start text
            if(writeDeletedText)
                _writer.WriteStartElement("w", "delText", OpenXmlNamespaces.WordprocessingML);
            else
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
                    //write page break, section breaks are written by writeParagraph() method
                    if (!isSectionEnd(cp))
                    {
                        _writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "page");
                        _writer.WriteEndElement();
                    }
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
                    _writer.WriteChars(new char[] { c }, 0, 1);

                }

                cp++;
            }

            //end text
            _writer.WriteEndElement();

            return cp;
        }

        #endregion

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
        /// Checks if the PAPX is old
        /// </summary>
        /// <param name="chpx">The PAPX</param>
        /// <returns></returns>
        private bool isOld(ParagraphPropertyExceptions papx)
        {
            bool ret = false;
            foreach (SinglePropertyModifier sprm in papx.grpprl)
            {
                if (sprm.OpCode == 0x2664)
                {
                    //sHasOldProps
                    ret = true;
                    break;
                }
            }
            return ret;
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
        /// Looks into the section table to find out if this CP is the end of a section
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        private bool isSectionEnd(Int32 cp)
        {
            bool result = false;

            //if cp is the last char of a section, the next section will start at cp +1
            int search = cp + 1;

            for (int i = 0; i < _doc.SectionTable.rgfc.Length; i++)
            {
                if (_doc.SectionTable.rgfc[i] == search)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        private Int32 getChpxRsid(CharacterPropertyExceptions chpx)
        {
            Int32 ret = 0;
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                if (sprm.OpCode == 0x6815 || 
                    sprm.OpCode == 0x6816 ||
                    sprm.OpCode == 0x6817)
                {
                    ret = System.BitConverter.ToInt32(sprm.Arguments, 0);
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

            if(_allPapx.ContainsKey(fc))
            {
                ret = _allPapx[fc];
                _lastValidPapx = ret;
            }
            else
            {
                ret = _lastValidPapx;
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
                _lastValidSepx = ret;
            }
            catch (KeyNotFoundException)
            {
                //there is no SEPX at this position, 
                //so the previous SEPX is valid for this cp

                Int32 lastKey = _doc.SectionTable.rgfc[1];
                foreach (Int32 key in _allSepx.Keys)
                {
                    if (cp > lastKey && cp < key)
                    {
                        ret = _allSepx[lastKey];
                        break;
                    }
                    else
                    {
                        lastKey = key;
                    }
                }
            }

            return ret;
        }
    }
}
