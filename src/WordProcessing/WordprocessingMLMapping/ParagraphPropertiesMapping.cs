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
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class ParagraphPropertiesMapping : PropertiesMapping,
          IMapping<ParagraphPropertyExceptions>
    {
        private StyleSheet _styleSheet;
        private XmlElement _pPr;

        public ParagraphPropertiesMapping(XmlWriter writer, StyleSheet styles)
            : base(writer)
        {
            _styleSheet = styles;
            _pPr = _nodeFactory.CreateElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
        }

        public void Apply(ParagraphPropertyExceptions papx)
        {
            XmlElement ind = _nodeFactory.CreateElement("w", "ind", OpenXmlNamespaces.WordprocessingML);
            XmlElement numPr = _nodeFactory.CreateElement("w", "numPr", OpenXmlNamespaces.WordprocessingML);
            XmlElement pBdr = _nodeFactory.CreateElement("w", "pBdr", OpenXmlNamespaces.WordprocessingML);
            XmlElement spacing = _nodeFactory.CreateElement("w", "spacing", OpenXmlNamespaces.WordprocessingML);
            XmlElement jc = null;
            XmlElement shd = null;

            //append style id 
            XmlElement pStyle = _nodeFactory.CreateElement("w", "pStyle", OpenXmlNamespaces.WordprocessingML);
            XmlAttribute styleId = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
            styleId.Value = StyleSheetMapping.MakeStyleId(_styleSheet.Styles[papx.istd].xstzName);
            pStyle.Attributes.Append(styleId);
            _pPr.AppendChild(pStyle);

            foreach (SinglePropertyModifier sprm in papx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //attributes
                    case 0x6465:
                        XmlAttribute divId = _nodeFactory.CreateAttribute("w", "divId", OpenXmlNamespaces.WordprocessingML);
                        divId.Value = System.BitConverter.ToUInt32(sprm.Arguments, 0).ToString();
                        _pPr.Attributes.Append(divId);
                        break;
                    case 0x2437:
                        appendFlagAttribute(_pPr, sprm, "autoSpaceDE");
                        break;
                    case 0x2438:
                        appendFlagAttribute(_pPr, sprm, "autoSpaceDN");
                        break;
                    case 0x2441:
                        appendFlagAttribute(_pPr, sprm, "bidi");
                        break;
                    case 0x246D:
                        appendFlagAttribute(_pPr, sprm, "contextualSpacing");
                        break;
                    
                    //element flags
                    case 0x2405:
                        appendFlagElement(_pPr, sprm, "keepLines");
                        break;
                    case 0x2406:
                        appendFlagElement(_pPr, sprm, "keepNext");
                        break;
                    case 0x2433:
                        appendFlagElement(_pPr, sprm, "kinsoku");
                        break;
                    case 0x2435:
                        appendFlagElement(_pPr, sprm, "overflowPunct");
                        break;
                    case 0x2407:
                        appendFlagElement(_pPr, sprm, "pageBreakBefore");
                        break;
                    case 0x242A:
                        appendFlagElement(_pPr, sprm, "su_pPressAutoHyphens");
                        break;
                    case 0x240C:
                        appendFlagElement(_pPr, sprm, "su_pPressLineNumbers");
                        break;
                    case 0x2462:
                        appendFlagElement(_pPr, sprm, "su_pPressOverlap");
                        break;
                    case 0x2436:
                        appendFlagElement(_pPr, sprm, "topLinePunct");
                        break;
                    case 0x2431:
                        appendFlagElement(_pPr, sprm, "widowControl");
                        break;    
                    case 0x2434:
                        appendFlagElement(_pPr, sprm, "wordWrap");
                        break;

                    //indentation
                    case 0x845e:
                    case 0x840F:
                    case 0x465F:
                    case 0x4610:
                        XmlAttribute left = _nodeFactory.CreateAttribute("w", "left", OpenXmlNamespaces.WordprocessingML);
                        left.Value = System.BitConverter.ToInt16(sprm.Arguments, 0).ToString();
                        ind.Attributes.Append(left);
                        break;
                    case 0x4456:
                        XmlAttribute leftChars = _nodeFactory.CreateAttribute("w", "leftChars", OpenXmlNamespaces.WordprocessingML);
                        leftChars.Value = System.BitConverter.ToInt16(sprm.Arguments, 0).ToString();
                        ind.Attributes.Append(leftChars);
                        break;
                    case 0x8460:
                    case 0x8411:
                        XmlAttribute firstLine = _nodeFactory.CreateAttribute("w", "firstLine", OpenXmlNamespaces.WordprocessingML);
                        firstLine.Value = System.BitConverter.ToInt16(sprm.Arguments, 0).ToString();
                        ind.Attributes.Append(firstLine);
                        break;
                    case 0x4457:
                        XmlAttribute firstLineChars = _nodeFactory.CreateAttribute("w", "firstLineChars", OpenXmlNamespaces.WordprocessingML);
                        firstLineChars.Value = System.BitConverter.ToInt16(sprm.Arguments, 0).ToString();
                        ind.Attributes.Append(firstLineChars);
                        break;
                    case 0x845D:
                    case 0x840E:
                        XmlAttribute right = _nodeFactory.CreateAttribute("w", "right", OpenXmlNamespaces.WordprocessingML);
                        right.Value = System.BitConverter.ToInt16(sprm.Arguments, 0).ToString();
                        ind.Attributes.Append(right);
                        break;
                    case 0x4455:
                        XmlAttribute rightChars = _nodeFactory.CreateAttribute("w", "rightChars", OpenXmlNamespaces.WordprocessingML);
                        rightChars.Value = System.BitConverter.ToInt16(sprm.Arguments, 0).ToString();
                        ind.Attributes.Append(rightChars);
                        break;

                    //spacing
                    case 0xA413:
                        XmlAttribute before = _nodeFactory.CreateAttribute("w", "before", OpenXmlNamespaces.WordprocessingML);
                        before.Value = System.BitConverter.ToUInt16(sprm.Arguments, 0).ToString();
                        spacing.Attributes.Append(before);
                        break;
                    case 0xA414:
                        XmlAttribute after = _nodeFactory.CreateAttribute("w", "after", OpenXmlNamespaces.WordprocessingML);
                        after.Value = System.BitConverter.ToUInt16(sprm.Arguments, 0).ToString();
                        spacing.Attributes.Append(after);
                        break;
                    case 0x245C:
                        XmlAttribute afterAutospacing = _nodeFactory.CreateAttribute("w", "afterAutospacing", OpenXmlNamespaces.WordprocessingML);
                        afterAutospacing.Value = sprm.Arguments[0].ToString();
                        spacing.Attributes.Append(afterAutospacing);
                        break;
                    case 0x245B:
                        XmlAttribute beforeAutospacing = _nodeFactory.CreateAttribute("w", "beforeAutospacing", OpenXmlNamespaces.WordprocessingML);
                        beforeAutospacing.Value = sprm.Arguments[0].ToString();
                        spacing.Attributes.Append(beforeAutospacing);
                        break;
                    case 0x6412:
                        LineSpacingDescriptor lspd = new LineSpacingDescriptor(sprm.Arguments);
                        XmlAttribute line = _nodeFactory.CreateAttribute("w", "line", OpenXmlNamespaces.WordprocessingML);
                        line.Value = Math.Abs(lspd.dyaLine).ToString();
                        spacing.Attributes.Append(line);
                        XmlAttribute lineRule = _nodeFactory.CreateAttribute("w", "lineRule", OpenXmlNamespaces.WordprocessingML);
                        if(!lspd.fMultLinespace && lspd.dyaLine < 0)
                            lineRule.Value = "exact";
                        else if(!lspd.fMultLinespace && lspd.dyaLine > 0)
                            lineRule.Value = "atLeast";
                        //no line rule means auto
                        spacing.Attributes.Append(lineRule);
                        break;

                    //justification code
                    case 0x2461:
                    case 0x2403:
                        jc = _nodeFactory.CreateElement("w", "jc", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute jcVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        jcVal.Value = ((Global.JustificationCode)sprm.Arguments[0]).ToString();
                        jc.Attributes.Append(jcVal);
                        break;

                    //numbering properties
                    case 0x260A:
                        XmlNode ilvl = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "ilvl", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute iLvlVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        iLvlVal.Value = sprm.Arguments[0].ToString();
                        ilvl.Attributes.Append(iLvlVal);
                        numPr.AppendChild(ilvl);
                        break;

                    //borders
                    case 0x461C:
                    case 0xC64E:
                    case 0x4424:
                    case 0x6424:
                        XmlNode topBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, topBorder);
                        addOrSetBorder(pBdr, topBorder);
                        break;
                    case 0x461D:
                    case 0xC64F:
                    case 0x4425:
                    case 0x6425:
                        XmlNode leftBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, leftBorder);
                        addOrSetBorder(pBdr, leftBorder);
                        break;
                    case 0x461E:
                    case 0xC650:
                    case 0x4426:
                    case 0x6426:
                        XmlNode bottomBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, bottomBorder);
                        addOrSetBorder(pBdr, bottomBorder);
                        break;
                    case 0x461F:
                    case 0xC651:
                    case 0x4427:
                    case 0x6427:
                        XmlNode rightBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, rightBorder);
                        addOrSetBorder(pBdr, rightBorder);
                        break;
                    case 0x4620:
                    case 0xC652:
                    case 0x4428:
                    case 0x6428:
                        XmlNode betweenBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "between", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, betweenBorder);
                        addOrSetBorder(pBdr, betweenBorder);
                        break;
                    case 0x4621:
                    case 0xC653:
                    case 0x4629:
                    case 0x6629:
                        XmlNode barBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bar", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(sprm.Arguments, barBorder);
                        addOrSetBorder(pBdr, barBorder);
                        break;
                    
                    //shading
                    case 0x442D:
                    case 0xc64d:
                        shd = _nodeFactory.CreateElement("w", "shd", OpenXmlNamespaces.WordprocessingML);
                        ShadingDescriptor desc = new ShadingDescriptor(sprm.Arguments);

                        //fill color
                        XmlAttribute fill = _nodeFactory.CreateAttribute("w", "fill", OpenXmlNamespaces.WordprocessingML);
                        if (desc.cvBack != 0)
                            fill.Value = String.Format("{0:x6}", desc.cvBack);
                        else
                            fill.Value = desc.icoBack.ToString();
                        shd.Attributes.Append(fill);

                        //foreground color
                        XmlAttribute color = _nodeFactory.CreateAttribute("w", "color", OpenXmlNamespaces.WordprocessingML);
                        if (desc.cvFore != 0)
                            color.Value = String.Format("{0:x6}", desc.cvFore);
                        else
                            color.Value = desc.icoFore.ToString();
                        shd.Attributes.Append(color);

                        //pattern
                        XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        val.Value = getShadingPattern(desc);
                        shd.Attributes.Append(val);

                        break;
                    
                    default:
                        break;
                }
            }

            //append indent
            if (ind.Attributes.Count > 0)
                _pPr.AppendChild(ind);

            //append spacing
            if (spacing.Attributes.Count > 0)
                _pPr.AppendChild(spacing);

            //append justification
            if (jc != null)
                _pPr.AppendChild(jc);

            //append numPr
            if (numPr.ChildNodes.Count > 0)
                _pPr.AppendChild(numPr);

            //append borders
            if (pBdr.ChildNodes.Count > 0)
                _pPr.AppendChild(pBdr);

            //append shading
            if (shd != null)
                _pPr.AppendChild(shd);

            //write Properties
            if (_pPr.ChildNodes.Count > 0 || _pPr.Attributes.Count > 0)
            {
                _pPr.WriteTo(_writer);
            }
        }

        private string getShadingPattern(ShadingDescriptor shd)
        {
            string pattern = "";
            switch (shd.ipat)
            {
                case ShadingDescriptor.ShadingPattern.Automatic:
                    pattern = "clear";
                    break;
                case ShadingDescriptor.ShadingPattern.Solid:
                    pattern = "solid";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_5:
                    pattern = "pct5";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_10:
                    pattern = "pct10";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_20:
                    pattern = "pct20";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_25:
                    pattern = "pct25";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_30:
                    pattern = "pct30";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_40:
                    pattern = "pct40";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_50:
                    pattern = "pct50";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_60:
                    pattern = "pct60";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_70:
                    pattern = "pct70";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_75:
                    pattern = "pct75";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_80:
                    pattern = "pct80";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_90:
                    pattern = "pct90";
                    break;
                case ShadingDescriptor.ShadingPattern.DarkHorizontal:
                    pattern = "thinHorzStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.DarkVertical:
                    pattern = "thinVertStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.DarkForwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.DarkBackwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.DarkCross:
                    break;
                case ShadingDescriptor.ShadingPattern.DarkDiagonalCross:
                    pattern = "thinDiagCross";
                    break;
                case ShadingDescriptor.ShadingPattern.Horizontal:
                    pattern = "horzStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.Vertical:
                    pattern = "vertStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.ForwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.BackwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.Cross:
                    break;
                case ShadingDescriptor.ShadingPattern.DiagonalCross:
                    pattern = "diagCross";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_2_5:
                    pattern = "pct5";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_7_5:
                    pattern = "pct10";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_12_5:
                    pattern = "pct12";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_15:
                    pattern = "pct15";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_17_5:
                    pattern = "pct15";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_22_5:
                    pattern = "pct20";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_27_5:
                    pattern = "pct30";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_32_5:
                    pattern = "pct35";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_35:
                    pattern = "pct35";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_37_5:
                    pattern = "pct37";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_42_5:
                    pattern = "pct40";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_45:
                    pattern = "pct45";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_47_5:
                    pattern = "pct45";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_52_5:
                    pattern = "pct50";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_55:
                    pattern = "pct55";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_57_5:
                    pattern = "pct55";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_62_5:
                    pattern = "pct62";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_65:
                    pattern = "pct65";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_67_5:
                    pattern = "pct65";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_72_5:
                    pattern = "pct70";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_77_5:
                    pattern = "pct75";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_82_5:
                    pattern = "pct80";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_85:
                    pattern = "pct85";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_87_5:
                    pattern = "pct87";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_92_5:
                    pattern = "pct90";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_95:
                    pattern = "pct95";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_97_5:
                    pattern = "pct95";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_97:
                    pattern = "pct95";
                    break;
                default:
                    pattern = "nil";
                    break;
            }
            return pattern;
        }

    }
}
