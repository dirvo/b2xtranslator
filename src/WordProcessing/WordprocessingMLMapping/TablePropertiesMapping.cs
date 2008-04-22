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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TablePropertiesMapping :
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private XmlElement _tblPr;
        private XmlElement _tblGrid;
        private StyleSheet _styles;
        private List<Int16> _grid;

        private enum WidthType
        {
            nil,
            auto,
            pct,
            dxa
        }

        private enum VerticalPositionCode
        {
            margin = 0,
            page,
            text
        }

        private enum HorizontalPositionCode
        {
            text = 0,
            margin,
            page
        }

        public TablePropertiesMapping(XmlWriter writer, StyleSheet styles, List<Int16> grid)
            : base(writer)
        {
            _styles = styles;
            _tblPr = _nodeFactory.CreateElement("w", "tblPr", OpenXmlNamespaces.WordprocessingML);
            _grid = grid;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            XmlElement tblBorders = _nodeFactory.CreateElement("w", "tblBorders", OpenXmlNamespaces.WordprocessingML);
            XmlElement tblCellMar = _nodeFactory.CreateElement("w", "tblCellMar", OpenXmlNamespaces.WordprocessingML);
            XmlElement tblLayout = _nodeFactory.CreateElement("w", "tblLayout", OpenXmlNamespaces.WordprocessingML);
            XmlElement tblpPr = _nodeFactory.CreateElement("w", "tblpPr", OpenXmlNamespaces.WordprocessingML);
            XmlAttribute layoutType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);

            layoutType.Value = "fixed";

            foreach (SinglePropertyModifier sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //preferred table width
                    case 0xF614:
                        WidthType fts = (WidthType)sprm.Arguments[0];
                        Int16 width = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        XmlElement tblW = _nodeFactory.CreateElement("w", "tblW", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute w = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                        w.Value = width.ToString();
                        XmlAttribute type = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                        type.Value = fts.ToString();
                        tblW.Attributes.Append(type);
                        tblW.Attributes.Append(w);
                        _tblPr.AppendChild(tblW);
                        break;

                    //justification
                    case 0x5400:
                    case 0x548A:
                        appendValueElement(_tblPr, "jc", ((Global.JustificationCode)sprm.Arguments[0]).ToString(), true);
                        break;

                    //indent
                    case 0xF661:
                        Int16 indValue = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        if (indValue != 0)
                        {
                            XmlElement tblInd = _nodeFactory.CreateElement("w", "tblInd", OpenXmlNamespaces.WordprocessingML);
                            XmlAttribute tblIndW = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                            tblIndW.Value = System.BitConverter.ToInt16(sprm.Arguments, 1).ToString();
                            tblInd.Attributes.Append(tblIndW);
                            XmlAttribute tblIndType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                            tblIndType.Value = "dxa";
                            tblInd.Attributes.Append(tblIndType);
                            _tblPr.AppendChild(tblInd);
                        }
                        break;

                    //style
                    case 0x563a:
                    case 0xd63d:
                        appendValueElement(_tblPr, "tblStyle", StyleSheetMapping.MakeStyleId(_styles.Styles[System.BitConverter.ToInt16(sprm.Arguments, 0)].xstzName), true);
                        break;

                    //bidi
                    case 0x560B:
                        appendValueElement(_tblPr, "bidiVisual", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //table look
                    case 0x740A:
                        appendValueElement(_tblPr, "tblLook", String.Format("{0:x4}", System.BitConverter.ToInt16(sprm.Arguments, 2)), true);
                        break;

                    //autofit
                    case 0x3615:
                        if (sprm.Arguments[0] == 1)
                            layoutType.Value = "auto";
                        break;

                    //default cell padding (margin)
                    case 0xD632:
                    case 0xD634:
                    case 0xD638:
                        byte grfbrc = sprm.Arguments[2];
                        Int16 wMar = System.BitConverter.ToInt16(sprm.Arguments, 4);
                        if (Utils.BitmaskToBool((int)grfbrc, 0x01))
                            appendDxaElement(tblCellMar, "top", wMar.ToString(), true);
                        if (Utils.BitmaskToBool((int)grfbrc, 0x02))
                            appendDxaElement(tblCellMar, "left", wMar.ToString(), true);
                        if (Utils.BitmaskToBool((int)grfbrc, 0x04))
                            appendDxaElement(tblCellMar, "bottom", wMar.ToString(), true);
                        if (Utils.BitmaskToBool((int)grfbrc, 0x08))
                            appendDxaElement(tblCellMar, "right", wMar.ToString(), true);
                        break;
                    
                    //default cell spacing
                    case 0xD631:
                    case 0xD633:
                    case 0xD637:
                        break;

                    //row count
                    case 0x3488:
                        appendValueElement(_tblPr, "tblStyleRowBandSize", sprm.Arguments[0].ToString(), true);
                        break;

                    //col count
                    case 0x3489:
                        appendValueElement(_tblPr, "tblStyleColBandSize", sprm.Arguments[0].ToString(), true);
                        break;

                    //overlap
                    case 0x3465:
                        bool noOverlap = Utils.ByteToBool(sprm.Arguments[0]);
                        string tblOverlapVal = "overlap";
                        if (noOverlap)
                            tblOverlapVal = "never";
                        appendValueElement(_tblPr, "tblOverlap", tblOverlapVal, true);
                        break;

                    //shading
                    case 0xD660:
                        ShadingDescriptor desc = new ShadingDescriptor(sprm.Arguments);
                        appendShading(_tblPr, desc);
                        break;

                    //borders
                    case 0xD613:
                        byte[] brc = new byte[8];
                        //top border
                        Array.Copy(sprm.Arguments, 0, brc, 0, 8);
                        XmlNode topBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(brc), topBorder1);
                        addOrSetBorder(tblBorders, topBorder1);
                        //left
                        Array.Copy(sprm.Arguments, 8, brc, 0, 8);
                        XmlNode leftBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(brc), leftBorder1);
                        addOrSetBorder(tblBorders, leftBorder1);
                        //bottom
                        Array.Copy(sprm.Arguments, 16, brc, 0, 8);
                        XmlNode bottomBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(brc), bottomBorder1);
                        addOrSetBorder(tblBorders, bottomBorder1);
                        //right
                        Array.Copy(sprm.Arguments, 24, brc, 0, 8);
                        XmlNode rightBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(brc), rightBorder1);
                        addOrSetBorder(tblBorders, rightBorder1);
                        //inside H
                        Array.Copy(sprm.Arguments, 32, brc, 0, 8);
                        XmlNode insideHBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideH", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(brc), insideHBorder1);
                        addOrSetBorder(tblBorders, insideHBorder1);
                        //inside V
                        Array.Copy(sprm.Arguments, 40, brc, 0, 8);
                        XmlNode insideVBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideV", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(brc), insideVBorder1);
                        addOrSetBorder(tblBorders, insideVBorder1);
                        break;
                    case 0xD47F:
                        XmlNode topBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), topBorder);
                        addOrSetBorder(tblBorders, topBorder);
                        break;
                    case 0xD680:
                        XmlNode bottomBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bottomBorder);
                        addOrSetBorder(tblBorders, bottomBorder);
                        break;
                    case 0xD681:
                        XmlNode leftBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), leftBorder);
                        addOrSetBorder(tblBorders, leftBorder);
                        break;
                    case 0xD682:
                        XmlNode rightBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), rightBorder);
                        addOrSetBorder(tblBorders, rightBorder);
                        break;
                    case 0xD683:
                        XmlNode insideHBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideH", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), insideHBorder);
                        addOrSetBorder(tblBorders, insideHBorder);
                        break;
                    case 0xD684:
                        XmlNode insideVBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideV", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), insideVBorder);
                        addOrSetBorder(tblBorders, insideVBorder);
                        break;

                    //floating table properties
                    case 0x360D:
                        byte flag = sprm.Arguments[0];
                        VerticalPositionCode pcVert = (VerticalPositionCode)((flag & 0x30) >> 4);
                        HorizontalPositionCode pcHorz = (HorizontalPositionCode)((flag & 0xC0) >> 6);
                        appendValueAttribute(tblpPr, "horzAnchor", pcHorz.ToString());
                        appendValueAttribute(tblpPr, "vertAnchor", pcVert.ToString());
                        break;
                    case 0x9410:
                        appendValueAttribute(tblpPr, "leftFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x941e:
                        appendValueAttribute(tblpPr, "rightFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x9411:
                        appendValueAttribute(tblpPr, "topFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x941f:
                        appendValueAttribute(tblpPr, "bottomFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x940e:
                        appendValueAttribute(tblpPr, "tblpX", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x940f:
                        appendValueAttribute(tblpPr, "tblpY", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                }
            }

            //append floating props
            if (tblpPr.Attributes.Count > 0)
            {
                _tblPr.AppendChild(tblpPr);
            }

            //append borders
            if(tblBorders.ChildNodes.Count > 0)
            {
                _tblPr.AppendChild(tblBorders);
            }

            //append layout type
            tblLayout.Attributes.Append(layoutType);
            _tblPr.AppendChild(tblLayout);

            //append margins
            if (tblCellMar.ChildNodes.Count > 0)
            {
                _tblPr.AppendChild(tblCellMar);
            }

            //write Properties
            if (_tblPr.ChildNodes.Count > 0 || _tblPr.Attributes.Count > 0)
            {
                _tblPr.WriteTo(_writer);
            }

            //append the grid
            _tblGrid = _nodeFactory.CreateElement("w", "tblGrid", OpenXmlNamespaces.WordprocessingML);
            foreach (Int16 colW in _grid)
            {
                XmlElement gridCol = _nodeFactory.CreateElement("w", "gridCol", OpenXmlNamespaces.WordprocessingML);
                XmlAttribute gridColW = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                gridColW.Value = colW.ToString();
                gridCol.Attributes.Append(gridColW);
                _tblGrid.AppendChild(gridCol);
            }
            _tblGrid.WriteTo(_writer);
        }
    }
}
