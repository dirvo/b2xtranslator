/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Collections;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TableCellPropertiesMapping : 
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private int _cellIndex;
        private XmlElement _tcPr;
        private XmlElement _tcMar;
        private XmlElement _tcBorders;
        private List<Int16> _grid;
        private Int16 _width;

        private const byte VMEGRE_CONTINUE = 1;
        private const byte VMERGE_RESTART = 3;

        private enum VerticalCellAlignment
        {
            top,
            center,
            bottom
        }

        public TableCellPropertiesMapping(XmlWriter writer, int cellIndex, List<Int16> tableGrid)
            : base(writer)
        {
            _tcPr = _nodeFactory.CreateElement("w", "tcPr", OpenXmlNamespaces.WordprocessingML);
            _tcMar = _nodeFactory.CreateElement("w", "tcMar", OpenXmlNamespaces.WordprocessingML);
            _tcBorders = _nodeFactory.CreateElement("w", "tcBorders", OpenXmlNamespaces.WordprocessingML);
            _cellIndex = cellIndex;
            _grid = tableGrid;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            foreach (SinglePropertyModifier sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)
	            {
                    //width
                    case 0xD608:
                        Int16 boundary2 = System.BitConverter.ToInt16(sprm.Arguments, 1 + ((_cellIndex + 1) * 2));
                        Int16 boundary1 = System.BitConverter.ToInt16(sprm.Arguments, 1 + (_cellIndex * 2));
                        _width = (Int16)(boundary2 - boundary1);
                        appendDxaElement(_tcPr, "tcW", _width.ToString(), true);
                        break;
                    case 0xD635:
                        byte first = sprm.Arguments[0];
                        byte lim = sprm.Arguments[1];
                        byte ftsWidth = sprm.Arguments[2];
                        _width = System.BitConverter.ToInt16(sprm.Arguments, 3);
                        if (_cellIndex >= first && _cellIndex < lim)
                            appendDxaElement(_tcPr, "tcW", _width.ToString(), true);
                        break;

                    //margins
                    case 0xd632:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        byte ftsMargin = sprm.Arguments[3];
                        Int16 wMargin = System.BitConverter.ToInt16(sprm.Arguments, 4);
                        if (_cellIndex >= first && _cellIndex < lim)
                        {
                            BitArray borderBits = new BitArray(new byte[] { sprm.Arguments[2] });
                            if (borderBits[0] == true)
                                appendDxaElement(_tcMar, "top", wMargin.ToString(), true);
                            if (borderBits[1] == true)
                                appendDxaElement(_tcMar, "left", wMargin.ToString(), true);
                            if (borderBits[2] == true)
                                appendDxaElement(_tcMar, "bottom", wMargin.ToString(), true);
                            if (borderBits[3] == true)
                                appendDxaElement(_tcMar, "right", wMargin.ToString(), true);
                        }
                        break;

                    //vertical alignment
                    case 0xD62C:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if(_cellIndex >= first && _cellIndex <lim)
                            appendValueElement(_tcPr, "vAlign", ((VerticalCellAlignment)sprm.Arguments[2]).ToString(), true);
                        break;

                    //vertical merge
                    case 0xD62B:
                        byte cell = sprm.Arguments[0];
                        byte mergeValue = sprm.Arguments[1];
                        if (cell == _cellIndex && mergeValue > 0)
                        {
                            switch (mergeValue)
                            {
                                case VMEGRE_CONTINUE:
                                    appendValueElement(_tcPr, "vMerge", "continue", true);
                                    break;
                                case VMERGE_RESTART:
                                    appendValueElement(_tcPr, "vMerge", "restart", true);
                                    break;
                            } 
                        }
                        break;

                    //Autofit
                    case 0xF636:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (_cellIndex >= first && _cellIndex < lim)
                            appendValueElement(_tcPr, "tcFitText", sprm.Arguments[2].ToString(), true);
                        break;

                    //no wrap
                    case 0xD639:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (_cellIndex >= first && _cellIndex < lim)
                            appendValueElement(_tcPr, "noWrap", sprm.Arguments[2].ToString(), true);
                        break;

                    //shading
                    case 0xD612:
                        //cell shading for cells 0-20
                        apppendCellShading(sprm.Arguments, _cellIndex);
                        break;
                    case 0xD616:
                        //cell shading for cells 21-42
                        apppendCellShading(sprm.Arguments, _cellIndex - 21);
                        break;
                    case 0xD60C:
                        //cell shading for cells 43-62
                        apppendCellShading(sprm.Arguments, _cellIndex - 43);
                        break;

                    //borders
                    case 0xD620:
                    case 0xD62F:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (_cellIndex >= first && _cellIndex < lim)
                        {
                            BitArray borderBits = new BitArray(new byte[] { sprm.Arguments[2] });
                            byte[] brc = new byte[sprm.Arguments.Length - 3];
                            Array.Copy(sprm.Arguments, 3, brc, 0, brc.Length);

                            if(borderBits[0])
                            {
                                //top border
                                XmlNode topBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                                appendBorderAttributes(brc, topBorder);
                                addOrSetBorder(_tcBorders, topBorder);
                            }
                            if (borderBits[1])
                            {
                                //left border
                                XmlNode leftBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                                appendBorderAttributes(brc, leftBorder);
                                addOrSetBorder(_tcBorders, leftBorder);
                            }
                            if (borderBits[2])
                            {
                                //bottom border
                                XmlNode bottomBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                                appendBorderAttributes(brc, bottomBorder);
                                addOrSetBorder(_tcBorders, bottomBorder);
                            }
                            if (borderBits[3])
                            {
                                //right border
                                XmlNode rightBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                                appendBorderAttributes(brc, rightBorder);
                                addOrSetBorder(_tcBorders, rightBorder);
                            }
                        }
                        break;
	            }
            }

            //horizontal merge
            if (_width > _grid[_cellIndex])
            {
                int span = 1;

                //check the number of merged cells
                int w = _grid[_cellIndex];
                for (int i = _cellIndex+1; i < _grid.Count; i++)
                {
                    span++;
                    w += _grid[i];
                    if (w == _width)
                        break;
                }

                appendValueElement(_tcPr, "gridSpan", span.ToString(), true);
            }

            //append margins
            if (_tcMar.ChildNodes.Count > 0)
            {
                _tcPr.AppendChild(_tcMar);
            }

            //append borders
            if (_tcBorders.ChildNodes.Count > 0)
            {
                _tcPr.AppendChild(_tcBorders);
            }

            //write Properties
            if (_tcPr.ChildNodes.Count > 0 || _tcPr.Attributes.Count > 0)
            {
                _tcPr.WriteTo(_writer);
            }
        }

        private void apppendCellShading(byte[] sprmArg, int cellIndex)
        {
            //shading descriptor can have 10 bytes (Word 2000) or 2 bytes (Word 97)
            int shdLength = 2;
            if (sprmArg.Length % 10 == 0)
                shdLength = 10;

            byte[] shdBytes = new byte[shdLength];

            //multiple cell can be formatted with the same SHD.
            //in this case there is only 1 SHD for all cells in the row.
            if ((cellIndex * shdLength) >= sprmArg.Length)
            {
                //use the first SHD
                cellIndex = 0;
            }

            Array.Copy(sprmArg, cellIndex * shdBytes.Length, shdBytes, 0, shdBytes.Length);
            
            ShadingDescriptor shd = new ShadingDescriptor(shdBytes);
            appendShading(_tcPr, shd);
        }
    }
}
