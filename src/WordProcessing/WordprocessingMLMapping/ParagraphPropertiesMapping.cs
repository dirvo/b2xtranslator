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
        private ConversionContext _ctx;
        private XmlElement _pPr;
        private SectionPropertyExceptions _sepx;
        private CharacterPropertyExceptions _paraEndChpx;
        private int _sectionNr;

        public ParagraphPropertiesMapping(XmlWriter writer, ConversionContext ctx, CharacterPropertyExceptions paraEndChpx)
            : base(writer)
        {
            _pPr = _nodeFactory.CreateElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
            _paraEndChpx = paraEndChpx;
            _ctx = ctx;
        }

        public ParagraphPropertiesMapping(
            XmlWriter writer, 
            ConversionContext ctx,
            CharacterPropertyExceptions paraEndChpx, 
            SectionPropertyExceptions sepx,
            int sectionNr)
            : base(writer)
        {
            _pPr = _nodeFactory.CreateElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
            _paraEndChpx = paraEndChpx;
            _sepx = sepx;
            _ctx = ctx;
            _sectionNr = sectionNr;
        }

        public void Apply(ParagraphPropertyExceptions papx)
        {
            XmlElement ind = _nodeFactory.CreateElement("w", "ind", OpenXmlNamespaces.WordprocessingML);
            XmlElement numPr = _nodeFactory.CreateElement("w", "numPr", OpenXmlNamespaces.WordprocessingML);
            XmlElement pBdr = _nodeFactory.CreateElement("w", "pBdr", OpenXmlNamespaces.WordprocessingML);
            XmlElement spacing = _nodeFactory.CreateElement("w", "spacing", OpenXmlNamespaces.WordprocessingML);
            XmlElement jc = null;

            //append style id , do not append "Normal" style (istd 0)
            if (papx.istd != 0)
            {
                XmlElement pStyle = _nodeFactory.CreateElement("w", "pStyle", OpenXmlNamespaces.WordprocessingML);
                XmlAttribute styleId = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                styleId.Value = StyleSheetMapping.MakeStyleId(_ctx.Doc.Styles.Styles[papx.istd].xstzName);
                pStyle.Attributes.Append(styleId);
                _pPr.AppendChild(pStyle);
            }

            //append formatting of paragraph end mark
            if (_paraEndChpx != null)
            {
                XmlElement rPr = _nodeFactory.CreateElement("w", "rPr", OpenXmlNamespaces.WordprocessingML);
                
                //append properties
                _paraEndChpx.Convert(new CharacterPropertiesMapping(rPr, _ctx.Doc, new RevisionData(_paraEndChpx)));


                RevisionData rev = new RevisionData(_paraEndChpx);
                //append delete infos
                if (rev.Type == RevisionData.RevisionType.Deleted)
                {
                    XmlElement del = _nodeFactory.CreateElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                    rPr.AppendChild(del);
                }

                if(rPr.ChildNodes.Count >0 )
                {
                    _pPr.AppendChild(rPr);
                }
            }

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
                        appendFlagElement(_pPr, sprm, "keepLines", true);
                        break;
                    case 0x2406:
                        appendFlagElement(_pPr, sprm, "keepNext", true);
                        break;
                    case 0x2433:
                        appendFlagElement(_pPr, sprm, "kinsoku", true);
                        break;
                    case 0x2435:
                        appendFlagElement(_pPr, sprm, "overflowPunct", true);
                        break;
                    case 0x2407:
                        appendFlagElement(_pPr, sprm, "pageBreakBefore", true);
                        break;
                    case 0x242A:
                        appendFlagElement(_pPr, sprm, "su_pPressAutoHyphens", true);
                        break;
                    case 0x240C:
                        appendFlagElement(_pPr, sprm, "su_pPressLineNumbers", true);
                        break;
                    case 0x2462:
                        appendFlagElement(_pPr, sprm, "su_pPressOverlap", true);
                        break;
                    case 0x2436:
                        appendFlagElement(_pPr, sprm, "topLinePunct", true);
                        break;
                    case 0x2431:
                        appendFlagElement(_pPr, sprm, "widowControl", true);
                        break;    
                    case 0x2434:
                        appendFlagElement(_pPr, sprm, "wordWrap", true);
                        break;

                    //indentation
                    case 0x845e:
                    case 0x840F:
                    case 0x465F:
                    case 0x4610:
                        appendValueAttribute(ind, "left", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x4456:
                        appendValueAttribute(ind, "leftChars", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x8460:
                    case 0x8411:
                        Int16 flValue = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        string flName;
                        if (flValue >= 0)
                        {
                            flName = "firstLine";
                        }
                        else
                        {
                            flName = "hanging";
                            flValue *= -1;
                        }
                        appendValueAttribute(ind, flName, flValue.ToString());
                        break;
                    case 0x4457:
                        appendValueAttribute(ind, "firstLineChars", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x845D:
                    case 0x840E:
                        appendValueAttribute(ind, "right", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x4455:
                        appendValueAttribute(ind, "rightChars", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
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

                    //borders
                    case 0x461C:
                    case 0xC64E:
                    case 0x4424:
                    case 0x6424:
                        XmlNode topBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), topBorder);
                        addOrSetBorder(pBdr, topBorder);
                        break;
                    case 0x461D:
                    case 0xC64F:
                    case 0x4425:
                    case 0x6425:
                        XmlNode leftBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), leftBorder);
                        addOrSetBorder(pBdr, leftBorder);
                        break;
                    case 0x461E:
                    case 0xC650:
                    case 0x4426:
                    case 0x6426:
                        XmlNode bottomBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bottomBorder);
                        addOrSetBorder(pBdr, bottomBorder);
                        break;
                    case 0x461F:
                    case 0xC651:
                    case 0x4427:
                    case 0x6427:
                        XmlNode rightBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), rightBorder);
                        addOrSetBorder(pBdr, rightBorder);
                        break;
                    case 0x4620:
                    case 0xC652:
                    case 0x4428:
                    case 0x6428:
                        XmlNode betweenBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "between", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), betweenBorder);
                        addOrSetBorder(pBdr, betweenBorder);
                        break;
                    case 0x4621:
                    case 0xC653:
                    case 0x4629:
                    case 0x6629:
                        XmlNode barBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bar", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), barBorder);
                        addOrSetBorder(pBdr, barBorder);
                        break;
                    
                    //shading
                    case 0x442D:
                    case 0xc64d:
                        ShadingDescriptor desc = new ShadingDescriptor(sprm.Arguments);
                        appendShading(_pPr, desc);
                        break;
                    
                    //numbering
                    case 0x260A:
                        appendValueElement(numPr, "ilvl", sprm.Arguments[0].ToString(), true);
                        break;
                    case 0x460B:
                        appendValueElement(numPr, "numId", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //tabs
                    case 0xc60D:
                    case 0xC615:
                        XmlElement tabs = _nodeFactory.CreateElement("w", "tabs", OpenXmlNamespaces.WordprocessingML);
                        int pos = 0;
                        //read the removed tabs
                        byte itbdDelMax = sprm.Arguments[pos];
                        pos++;
                        for(int i=0; i<itbdDelMax; i++)
                        {
                            XmlElement tab = _nodeFactory.CreateElement("w", "tab", OpenXmlNamespaces.WordprocessingML);
                            //clear
                            XmlAttribute tabsVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                            tabsVal.Value = "clear";
                            tab.Attributes.Append(tabsVal);
                            //position
                            XmlAttribute tabsPos = _nodeFactory.CreateAttribute("w", "pos", OpenXmlNamespaces.WordprocessingML);
                            tabsPos.Value = System.BitConverter.ToInt16(sprm.Arguments, pos).ToString();
                            tab.Attributes.Append(tabsPos);
                            tabs.AppendChild(tab);
                            
                            //skip the tolerence array in sprm 0xC615
                            if (sprm.OpCode == 0xC615)
                                pos += 4;
                            else
                                pos += 2;
                        }
                        //read the added tabs
                        byte itbdAddMax = sprm.Arguments[pos];
                        pos++;
                        for (int i = 0; i < itbdAddMax; i++)
                        {
                            TabDescriptor tbd = new TabDescriptor(sprm.Arguments[pos + (itbdAddMax * 2) + i]);
                            XmlElement tab = _nodeFactory.CreateElement("w", "tab", OpenXmlNamespaces.WordprocessingML);
                            //justification
                            XmlAttribute tabsVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                            tabsVal.Value = ((Global.JustificationCode)tbd.jc).ToString();
                            tab.Attributes.Append(tabsVal);
                            //tab leader type
                            XmlAttribute leader = _nodeFactory.CreateAttribute("w", "leader", OpenXmlNamespaces.WordprocessingML);
                            leader.Value = ((Global.TabLeader)tbd.tlc).ToString();
                            tab.Attributes.Append(leader);
                            //position
                            XmlAttribute tabsPos = _nodeFactory.CreateAttribute("w", "pos", OpenXmlNamespaces.WordprocessingML);
                            tabsPos.Value = System.BitConverter.ToInt16(sprm.Arguments, pos + (i * 2)).ToString();
                            tab.Attributes.Append(tabsPos);
                            tabs.AppendChild(tab);
                        }
                        _pPr.AppendChild(tabs);
                        break;

                    default:
                        break;
                }
            }

            //append section properties
            if (_sepx != null)
            {
                XmlElement sectPr = _nodeFactory.CreateElement("w", "sectPr", OpenXmlNamespaces.WordprocessingML);
                _sepx.Convert(new SectionPropertiesMapping(sectPr, _ctx, _sectionNr));
                _pPr.AppendChild(sectPr);
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

            //write Properties
            if (_pPr.ChildNodes.Count > 0 || _pPr.Attributes.Count > 0)
            {
                _pPr.WriteTo(_writer);
            }
        }
    }
}
