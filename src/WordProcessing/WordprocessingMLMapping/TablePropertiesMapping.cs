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

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TablePropertiesMapping :
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private XmlElement _tblPr;
        private XmlElement _tblGrid;

        public TablePropertiesMapping(XmlWriter writer)
            : base(writer)
        {
            _tblPr = _nodeFactory.CreateElement("w", "tblPr", OpenXmlNamespaces.WordprocessingML);
            _tblGrid = _nodeFactory.CreateElement("w", "tblGrid", OpenXmlNamespaces.WordprocessingML);
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            XmlElement tblBorders = _nodeFactory.CreateElement("w", "tblBorders", OpenXmlNamespaces.WordprocessingML);
            XmlElement tblCellMar = _nodeFactory.CreateElement("w", "tblCellMar", OpenXmlNamespaces.WordprocessingML);

            foreach (SinglePropertyModifier sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //preferred table width
                    case 0xF614:
                        XmlElement tblW = _nodeFactory.CreateElement("w", "tblW", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute w = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                        w.Value = System.BitConverter.ToInt16(sprm.Arguments, 1).ToString();
                        XmlAttribute type = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                        type.Value = "dxa";
                        tblW.Attributes.Append(type);
                        tblW.Attributes.Append(w);
                        _tblPr.AppendChild(tblW);
                        break;

                    //autofit
                    case 0x3615:
                        XmlElement tblLayout = _nodeFactory.CreateElement("w", "tblLayout", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute layoutType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                        if (sprm.Arguments[0] == 1)
                            layoutType.Value = "auto";
                        else
                            layoutType.Value = "fixed";
                        tblLayout.Attributes.Append(layoutType);
                        _tblPr.AppendChild(tblLayout);
                        break;

                    //justification
                    case 0x5400:
                    case 0x548A:
                        appendValueElement(_tblPr, "jc", ((Global.JustificationCode)sprm.Arguments[0]).ToString());
                        break;

                    //cell padding
                    case 0xD632:
                    case 0xD634:
                    case 0xD638:
                        //not yet implemented
                        break;

                    //cell spacing
                    case 0xD631:
                    case 0xD633:
                    case 0xD637:
                        //not yet implemented
                        break;

                    //borders
                    case 0xD613:
                        byte[] brc = new byte[8];
                        //top border
                        Array.Copy(sprm.Arguments, 0, brc, 0, 8);
                        XmlNode topBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(brc, topBorder1);
                        addOrSetBorder(tblBorders, topBorder1);
                        //bottom
                        Array.Copy(sprm.Arguments, 16, brc, 0, 8);
                        XmlNode bottomBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(brc, bottomBorder1);
                        addOrSetBorder(tblBorders, bottomBorder1);
                        //left
                        Array.Copy(sprm.Arguments, 8, brc, 0, 8);
                        XmlNode leftBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(brc, leftBorder1);
                        addOrSetBorder(tblBorders, leftBorder1);
                        //right
                        Array.Copy(sprm.Arguments, 24, brc, 0, 8);
                        XmlNode rightBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(brc, rightBorder1);
                        addOrSetBorder(tblBorders, rightBorder1);
                        //inside H
                        Array.Copy(sprm.Arguments, 32, brc, 0, 8);
                        XmlNode insideHBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideH", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(brc, insideHBorder1);
                        addOrSetBorder(tblBorders, insideHBorder1);
                        //inside V
                        Array.Copy(sprm.Arguments, 40, brc, 0, 8);
                        XmlNode insideVBorder1 = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideV", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(brc, insideVBorder1);
                        addOrSetBorder(tblBorders, insideVBorder1);
                        break;
                    case 0xD47F:
                        XmlNode topBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, topBorder);
                        addOrSetBorder(tblBorders, topBorder);
                        break;
                    case 0xD680:
                        XmlNode bottomBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, bottomBorder);
                        addOrSetBorder(tblBorders, bottomBorder);
                        break;
                    case 0xD681:
                        XmlNode leftBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, leftBorder);
                        addOrSetBorder(tblBorders, leftBorder);
                        break;
                    case 0xD682:
                        XmlNode rightBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, rightBorder);
                        addOrSetBorder(tblBorders, rightBorder);
                        break;
                    case 0xD683:
                        XmlNode insideHBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideH", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, insideHBorder);
                        addOrSetBorder(tblBorders, insideHBorder);
                        break;
                    case 0xD684:
                        XmlNode insideVBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideV", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, insideVBorder);
                        addOrSetBorder(tblBorders, insideVBorder);
                        break;

                    default:
                        break;
                }
            }

            //append borders
            if(tblBorders.ChildNodes.Count > 0)
            {
                _tblPr.AppendChild(tblBorders);
            }

            //write Properties
            if (_tblPr.ChildNodes.Count > 0 || _tblPr.Attributes.Count > 0)
            {
                _tblPr.WriteTo(_writer);
            }
        }
    }
}
