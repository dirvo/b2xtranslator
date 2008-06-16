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
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO.Compression;
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.Threading;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public abstract class DocumentMapping : 
        AbstractOpenXmlMapping,
        IMapping<WordDocument>
    {
        protected WordDocument _doc;
        protected ConversionContext _ctx;
        protected ParagraphPropertyExceptions _lastValidPapx;
        protected SectionPropertyExceptions _lastValidSepx;
        protected int _sectionNr = 0;
        protected int _footnoteNr = 0;
        protected ContentPart _targetPart;

        private class Symbol
        {
            public string FontName;
            public string HexValue;
        }

        /// <summary>
        /// Creates a new DocumentMapping that writes to the given XmlWriter
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="targetPart"></param>
        public DocumentMapping(ConversionContext ctx, ContentPart targetPart, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
            _targetPart = targetPart;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        /// <summary>
        /// Creates a new DocumentMapping that creates a new XmLWriter on to the given ContentPart
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="targetPart"></param>
        public DocumentMapping(ConversionContext ctx, ContentPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), ctx.WriterSettings))
        {
            _ctx = ctx;
            _targetPart = targetPart;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        public abstract void Apply(WordDocument doc);

        #region TableConversion

        /// <summary>
        /// Writes the table starts at the given cp value
        /// </summary>
        /// <param name="cp">The cp at where the table begins</param>
        /// <returns>The character pointer to the first character after this table</returns>
        protected Int32 writeTable(Int32 initialCp)
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
        protected Int32 writeTableRow(Int32 initialCp, List<Int16> grid)
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
            while (!(_doc.Text[cp] == TextMark.CellOrRowMark && tai.fTtp) && tai.fInTable)
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
        protected List<Int16> buildTableGrid(int initialCp)
        {
            ParagraphPropertyExceptions backup = _lastValidPapx;

            List<Int16> boundaries = new List<Int16>();
            List<Int16> grid = new List<Int16>();
            Int32 cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo tai = new TableInfo(papx);

            Int32 fcRowEnd = findRowEndFc(cp, out cp);

            while (tai.fInTable)
            {
                //check all SPRMs of this TAPX
                foreach (SinglePropertyModifier sprm in papx.grpprl)
                {
                    //find the tDef SPRM
                    if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmTDefTable)
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
        protected Int32 findRowEndFc(int initialCp, out int rowEndCp)
        {
            int cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo tai = new TableInfo(papx);

            while (tai.fTtp==false && tai.fInTable==true)
            {
                while (_doc.Text[cp] != TextMark.CellOrRowMark)
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
        protected Int32 findRowEndFc(int initialCp)
        {
            int cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            ParagraphPropertyExceptions papx = findValidPapx(fc);
            TableInfo tai = new TableInfo(papx);

            while (tai.fTtp == false && tai.fInTable == true)
            {
                while (_doc.Text[cp] != TextMark.CellOrRowMark)
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
        protected Int32 writeTableCell(Int32 initialCp, TablePropertyExceptions tapx, List<Int16> grid, ref int gridIndex, int cellIndex)
        {
            Int32 cp = initialCp;

            //start w:tc
            _writer.WriteStartElement("w", "tc", OpenXmlNamespaces.WordprocessingML);

            //find cell end
            Int32 cpCellEnd = initialCp;
            while (_doc.Text[cpCellEnd] != TextMark.CellOrRowMark)
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
        protected Int32 writeParagraph(Int32 cp) 
        {
            //search the paragraph end
            Int32 cpParaEnd = cp;
            while (_doc.Text[cpParaEnd] != TextMark.ParagraphEnd && 
                _doc.Text[cpParaEnd] != TextMark.CellOrRowMark &&
                !(_doc.Text[cpParaEnd] == TextMark.PageBreakOrSectionMark && isSectionEnd(cpParaEnd) ))
            {
                cpParaEnd++;
            }

            if (_doc.Text[cpParaEnd] == TextMark.PageBreakOrSectionMark)
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
        protected Int32 writeParagraph(Int32 initialCp, Int32 cpEnd, bool sectionEnd)
        {
            Int32 cp = initialCp;
            Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
            Int32 fcEnd = _doc.PieceTable.FileCharacterPositions[cpEnd];
            ParagraphPropertyExceptions papx = findValidPapx(fc);

            //get all CHPX between these boundaries to determine the count of runs
            List<CharacterPropertyExceptions> chpxs = _doc.GetCharacterPropertyExceptions(fc, fcEnd);
            List<Int32> chpxFcs = _doc.GetFileCharacterPositions(fc, fcEnd);
            chpxFcs.Add(fcEnd);

            //the last of these CHPX formats the paragraph end mark
            CharacterPropertyExceptions paraEndChpx = chpxs[chpxs.Count-1];

            //start paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //check for section properties
            if (sectionEnd)
            {
                //this is the last paragraph of this section
                //write properties with section properties
                papx.Convert(new ParagraphPropertiesMapping(_writer, _ctx, paraEndChpx, findValidSepx(cpEnd), _sectionNr));
                _sectionNr++;
            }
            else
            {
                //write properties
                papx.Convert(new ParagraphPropertiesMapping(_writer, _ctx, paraEndChpx));
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
        protected Int32 writeRun(List<char> chars, CharacterPropertyExceptions chpx, Int32 initialCp)
        {
            Int32 cp = initialCp;

            RevisionData rev = new RevisionData(chpx);

            if (rev.Type == RevisionData.RevisionType.Deleted)
            {
                //If it's a deleted run
                _writer.WriteStartElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, "[b2x: could not retrieve author]");
                _writer.WriteAttributeString("w", "date", OpenXmlNamespaces.WordprocessingML, "[b2x: could not retrieve date]");
            }
            else if (rev.Type == RevisionData.RevisionType.Inserted)
            {
                //if it's a inserted run
                _writer.WriteStartElement("w", "ins", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, _doc.AuthorTable[rev.Isbt]);
                rev.Dttm.Convert(new DateMapping(_writer));
            }

            //start run
            _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);

            //append rsids
            if (rev.Rsid != 0)
            {
                string rsid = String.Format("{0:x8}", rev.Rsid);
                _writer.WriteAttributeString("w", "rsidR", OpenXmlNamespaces.WordprocessingML, rsid);
                _ctx.AddRsid(rsid);
            }
            if (rev.RsidDel != 0)
            {
                string rsidDel = String.Format("{0:x8}", rev.RsidDel);
                _writer.WriteAttributeString("w", "rsidDel", OpenXmlNamespaces.WordprocessingML, rsidDel);
                _ctx.AddRsid(rsidDel);
            }
            if(rev.RsidProp != 0)
            {
                string rsidProp = String.Format("{0:x8}", rev.RsidProp);
                _writer.WriteAttributeString("w", "rsidRPr", OpenXmlNamespaces.WordprocessingML, rsidProp);
                _ctx.AddRsid(rsidProp);
            }

            //convert properties
            chpx.Convert(new CharacterPropertiesMapping(_writer, _doc, rev, _lastValidPapx, false));

            if(rev.Type == RevisionData.RevisionType.Deleted)
                cp = writeText(chars, cp, chpx, true);
            else
                cp = writeText(chars, cp, chpx, false);

            //end run
            _writer.WriteEndElement();

            if (rev.Type == RevisionData.RevisionType.Deleted || rev.Type == RevisionData.RevisionType.Inserted)
            {
                _writer.WriteEndElement();
            }

            return cp;
        }

        /// <summary>
        /// Writes the given text to the document
        /// </summary>
        /// <param name="chars"></param>
        protected Int32 writeText(List<char> chars, Int32 initialCp, CharacterPropertyExceptions chpx, bool writeDeletedText)
        {
            Int32 cp = initialCp;
            bool fSpec = isSpecial(chpx);

            //start text
            string textType = "t";

            if(writeDeletedText)
                textType = "delText";

            _writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);

            if ((int)chars[0] == 32 || (int)chars[chars.Count - 1] == 32)
            {
                _writer.WriteAttributeString("xml", "space", "", "preserve");
            }

            //write text
            foreach (char c in chars)
            {
                if (c == TextMark.Tab)
                {
                    _writer.WriteEndElement();
                    _writer.WriteElementString("w", "tab", OpenXmlNamespaces.WordprocessingML, "");
                    _writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);
                }
                else if (c == TextMark.HardLineBreak)
                {
                    _writer.WriteElementString("w", "br", OpenXmlNamespaces.WordprocessingML, "");
                }
                else if (c == TextMark.ParagraphEnd)
                {
                    //do nothing
                }
                else if (c == TextMark.PageBreakOrSectionMark)
                {
                    //write page break, section breaks are written by writeParagraph() method
                    if (!isSectionEnd(cp))
                    {
                        _writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "page");
                        _writer.WriteEndElement();
                    }
                }
                else if (c == TextMark.ColumnBreak)
                {
                    _writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "column");
                    _writer.WriteEndElement();
                }
                else if (c == TextMark.FieldBeginMark)
                {
                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "begin");
                    _writer.WriteEndElement();
                }
                else if (c == TextMark.FieldSeperator)
                {
                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "separate");
                    _writer.WriteEndElement();
                }
                else if (c == TextMark.FieldEndMark)
                {
                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "end");
                    _writer.WriteEndElement();
                }
                else if(c == TextMark.Symbol && fSpec)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    Symbol s = getSymbol(chpx);
                    _writer.WriteStartElement("w", "sym", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "font", OpenXmlNamespaces.WordprocessingML, s.FontName);
                    _writer.WriteAttributeString("w", "char", OpenXmlNamespaces.WordprocessingML, s.HexValue);
                    _writer.WriteEndElement();

                    _writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);
                }
                else if (c == TextMark.DrawnObject && fSpec)
                {
                    FileShapeAddress fspa = null;
                    if(GetType() == typeof(MainDocumentMapping) && _doc.OfficeDrawingTable.ContainsKey(cp))
                    {
                         fspa = _doc.OfficeDrawingTable[cp];
                    }
                    else if(GetType() == typeof(HeaderMapping))
                    {
                        int headerCp = cp - _doc.FIB.ccpText - _doc.FIB.ccpFtn;
                        if (_doc.OfficeDrawingTableHeader.ContainsKey(headerCp))
                        {
                            fspa = _doc.OfficeDrawingTableHeader[headerCp];
                        }
                    }
                    if (fspa != null && fspa.ShapeContainer != null)
                    {
                        //close previous w:t ...
                        _writer.WriteEndElement();
                        fspa.ShapeContainer.Convert(new VMLShapeMapping(_writer, _targetPart, fspa, true, _ctx));
                        _writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);
                    }
                }
                else if (c == TextMark.Picture && fSpec)
                {
                    PictureDescriptor pict = new PictureDescriptor(chpx, _doc.DataStream);
                    if (pict.mfp.mm > 98 && pict.ShapeContainer != null)
                    {
                        //close previous w:t ...
                        _writer.WriteEndElement();
                        pict.Convert(new VMLPictureMapping(_writer, _targetPart));
                        _writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);
                    }
                }
                else if (c == TextMark.AutoNumberedFootnoteReference && fSpec)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    if (this.GetType() != typeof(FootnotesMapping))
                    {
                        _writer.WriteStartElement("w", "footnoteReference", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, _footnoteNr.ToString());
                        _writer.WriteEndElement();
                    }
                    else
                    {
                        _writer.WriteElementString("w", "footnoteRef", OpenXmlNamespaces.WordprocessingML, "");
                    }

                    _footnoteNr++;

                    _writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);
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

        #region HelpFunctions



        /// <summary>
        /// Checks if the PAPX is old
        /// </summary>
        /// <param name="chpx">The PAPX</param>
        /// <returns></returns>
        protected bool isOld(ParagraphPropertyExceptions papx)
        {
            bool ret = false;
            foreach (SinglePropertyModifier sprm in papx.grpprl)
            {
                if(sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPWall)
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
        protected bool isSpecial(CharacterPropertyExceptions chpx)
        {
            bool ret = false;
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCPicLocation ||
                    sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCHsp)
                {
                    //special picture
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCSymbol)
                {
                    //special symbol
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCFSpec)
                {
                    //special value
                    ret = Utils.ByteToBool(sprm.Arguments[0]);
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chpx"></param>
        /// <returns></returns>
        private Symbol getSymbol(CharacterPropertyExceptions chpx)
        {
            Symbol ret = null;
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCSymbol)
                {
                    //special symbol
                    ret = new Symbol();
                    Int16 fontIndex = System.BitConverter.ToInt16(sprm.Arguments, 0);
                    Int16 code = System.BitConverter.ToInt16(sprm.Arguments, 2);
                    ret.FontName = _doc.FontTable[fontIndex].xszFtn;
                    ret.HexValue = String.Format("{0:x4}", code);
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
        protected bool isSectionEnd(Int32 cp)
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

        /// <summary>
        /// Finds the PAPX that is valid for the given FC.
        /// </summary>
        /// <param name="fc"></param>
        /// <returns></returns>
        protected ParagraphPropertyExceptions findValidPapx(Int32 fc)
        {
            ParagraphPropertyExceptions ret = null;

            if(_ctx.AllPapx.ContainsKey(fc))
            {
                ret = _ctx.AllPapx[fc];
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
        protected SectionPropertyExceptions findValidSepx(Int32 cp)
        {
            SectionPropertyExceptions ret = null;

            try
            {
                ret = _ctx.AllSepx[cp];
                _lastValidSepx = ret;
            }
            catch (KeyNotFoundException)
            {
                //there is no SEPX at this position, 
                //so the previous SEPX is valid for this cp

                Int32 lastKey = _doc.SectionTable.rgfc[1];
                foreach (Int32 key in _ctx.AllSepx.Keys)
                {
                    if (cp > lastKey && cp < key)
                    {
                        ret = _ctx.AllSepx[lastKey];
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

        #endregion
    }
}
