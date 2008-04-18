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
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class CharacterPropertiesMapping : PropertiesMapping,
          IMapping<CharacterPropertyExceptions>
    {
        private WordDocument _doc;
        private XmlElement _rPr;
        private UInt16 _currentIstd;
        private RevisionData _revisionData;

        public CharacterPropertiesMapping(XmlWriter writer, WordDocument doc, RevisionData rev)
            : base(writer)
        {
            _doc = doc;
            _rPr = _nodeFactory.CreateElement("w", "rPr", OpenXmlNamespaces.WordprocessingML);
            _revisionData = rev;
        }

        public CharacterPropertiesMapping(XmlElement rPr, WordDocument doc, RevisionData rev)
            : base(null)
        {
            _doc = doc;
            _nodeFactory = rPr.OwnerDocument;
            _rPr = rPr;
            _revisionData = rev;
        }

        public void Apply(CharacterPropertyExceptions chpx)
        {

            //convert the normal SPRMS
            convertSprms(chpx.grpprl, _rPr);

            //apend revision changes
            if (_revisionData.Type == RevisionData.RevisionType.Changed)
            {
                XmlElement rPrChange = _nodeFactory.CreateElement("w", "rPrChange", OpenXmlNamespaces.WordprocessingML);

                //rsid
                XmlAttribute id = _nodeFactory.CreateAttribute("w", "id", OpenXmlNamespaces.WordprocessingML);
                id.Value = _revisionData.Rsid.ToString();
                rPrChange.Attributes.Append(id);

                //date
                _revisionData.Dttm.Convert(new DateMapping(rPrChange));

                //author
                XmlAttribute author = _nodeFactory.CreateAttribute("w", "author", OpenXmlNamespaces.WordprocessingML);
                author.Value = _doc.AuthorTable[_revisionData.Isbt];
                rPrChange.Attributes.Append(author);

                //convert revision stack
                convertSprms(_revisionData.Changes, rPrChange);

                _rPr.AppendChild(rPrChange);
            }

            //write properties
            if (_writer!=null && (_rPr.ChildNodes.Count > 0 || _rPr.Attributes.Count > 0))
            {
                _rPr.WriteTo(_writer);
            }
        }

        private void convertSprms(List<SinglePropertyModifier> sprms, XmlElement parent)
        {
            XmlElement shd = _nodeFactory.CreateElement("w", "shd", OpenXmlNamespaces.WordprocessingML);
            XmlElement rFonts = _nodeFactory.CreateElement("w", "rFonts", OpenXmlNamespaces.WordprocessingML);
            XmlElement color = _nodeFactory.CreateElement("w", "color", OpenXmlNamespaces.WordprocessingML);
            XmlAttribute colorVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
            XmlElement lang = _nodeFactory.CreateElement("w", "lang", OpenXmlNamespaces.WordprocessingML);

            foreach (SinglePropertyModifier sprm in sprms)
            {
                //no style is set at the moment
                _currentIstd = UInt16.MaxValue;

                switch (sprm.OpCode)
                {
                    //style id 
                    case 0x4A30:
                        _currentIstd = System.BitConverter.ToUInt16(sprm.Arguments, 0);
                        appendValueElement(parent, "rStyle", StyleSheetMapping.MakeStyleId(_doc.Styles.Styles[_currentIstd].xstzName), true);
                        break;
                    
                    //Element flags
                    case 0x085A:
                        appendFlagElement(parent, sprm, "rtl", true);
                        break;
                    case 0x0835:
                        appendFlagElement(parent, sprm, "b", true);
                        break;
                    case 0x085C:
                        appendFlagElement(parent, sprm, "bCs", true);
                        break;
                    case 0x083B:
                        appendFlagElement(parent, sprm, "caps", true); ;
                        break;
                    case 0x0882:
                        appendFlagElement(parent, sprm, "cs", true);
                        break;
                    case 0x2A53:
                        appendFlagElement(parent, sprm, "dstrike", true);
                        break;
                    case 0x0858:
                        appendFlagElement(parent, sprm, "emboss", true);
                        break;
                    case 0x0854:
                        appendFlagElement(parent, sprm, "imprint", true);
                        break;
                    case 0x0836:
                        appendFlagElement(parent, sprm, "i", true);
                        break;
                    case 0x085D:
                        appendFlagElement(parent, sprm, "iCs", true);
                        break;
                    case 0x0875:
                        appendFlagElement(parent, sprm, "noProof", true);
                        break;
                    case 0x0838:
                        appendFlagElement(parent, sprm, "outline", true);
                        break;
                    case 0x0839:
                        appendFlagElement(parent, sprm, "shadow", true);
                        break;
                    case 0x083A:
                        appendFlagElement(parent, sprm, "smallCaps", true);
                        break;
                    case 0x0818:
                        appendFlagElement(parent, sprm, "specVanish", true);
                        break;
                    case 0x0837:
                        appendFlagElement(parent, sprm, "strike", true);
                        break;
                    case 0x083C:
                        appendFlagElement(parent, sprm, "vanish", true);
                        break;
                    case 0x0811:
                        appendFlagElement(parent, sprm, "webHidden", true);
                        break;

                    //language
                    case 0x486D:
                    case 0x4873:
                        //latin
                        Int16 langid = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        if(langid != 1024)
                        {
                            XmlAttribute langVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                            langVal.Value = langid.ToString();
                            lang.Attributes.Append(langVal);
                        }
                        break;
                    case 0x486E:
                    case 0x4874:
                        //east asia
                        langid = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        if (langid != 1024)
                        {
                            XmlAttribute langEastAsia = _nodeFactory.CreateAttribute("w", "eastAsia", OpenXmlNamespaces.WordprocessingML);
                            langEastAsia.Value = langid.ToString();
                            lang.Attributes.Append(langEastAsia);
                        }
                        break;
                    case 0x485F:
                        //bidi
                        langid = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        if (langid != 1024)
                        {
                            XmlAttribute langBidi = _nodeFactory.CreateAttribute("w", "bidi", OpenXmlNamespaces.WordprocessingML);
                            langBidi.Value = langid.ToString();
                            lang.Attributes.Append(langBidi);
                        }
                        break;
                    
                    //borders
                    case 0x6865:
                    case 0xCA72:
                        XmlNode bdr = _nodeFactory.CreateElement("w", "bdr", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bdr);
                        parent.AppendChild(bdr);
                        break;

                    //shading
                    case 0x4866:
                    case 0xCA71:
                        ShadingDescriptor desc = new ShadingDescriptor(sprm.Arguments);
                        appendShading(parent, desc);
                        break;

                    //color
                    case 0x2A42:
                    case 0x4A60:
                        colorVal.Value = ((Global.ColorIdentifier)(sprm.Arguments[0])).ToString();
                        break;
                    case 0x6870:
                        //R
                        colorVal.Value = String.Format("{0:x2}", sprm.Arguments[0]);
                        //G
                        colorVal.Value += String.Format("{0:x2}", sprm.Arguments[1]);
                        //B
                        colorVal.Value += String.Format("{0:x2}", sprm.Arguments[2]);
                        break;

                    //highlightning
                    case 0x2A0C:
                        appendValueElement(parent, "highlight", ((Global.ColorIdentifier)sprm.Arguments[0]).ToString(), true);
                        break;

                    //spacing
                    case 0x8840:
                        appendValueElement(parent, "spacing", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //font size
                    case 0x4A43:
                        appendValueElement(parent, "sz", sprm.Arguments[0].ToString(), true);
                        break;
                    case 0x484B:
                        appendValueElement(parent, "kern", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;
                    case 0x4A61:
                        appendValueElement(parent, "szCs", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //font family
                    case 0x4A4F:
                       XmlAttribute ascii = _nodeFactory.CreateAttribute("w", "ascii", OpenXmlNamespaces.WordprocessingML);
                       ascii.Value = _doc.FontTable[System.BitConverter.ToUInt16(sprm.Arguments, 0)].xszFtn;
                       rFonts.Attributes.Append(ascii);
                       break;
                   case 0x4A50:
                       XmlAttribute eastAsia = _nodeFactory.CreateAttribute("w", "eastAsia", OpenXmlNamespaces.WordprocessingML);
                       eastAsia.Value = _doc.FontTable[System.BitConverter.ToUInt16(sprm.Arguments, 0)].xszFtn;
                       rFonts.Attributes.Append(eastAsia);
                       break;
                    case 0x4A51:
                        XmlAttribute ansi = _nodeFactory.CreateAttribute("w", "hAnsi", OpenXmlNamespaces.WordprocessingML);
                        ansi.Value = _doc.FontTable[System.BitConverter.ToUInt16(sprm.Arguments, 0)].xszFtn;
                        rFonts.Attributes.Append(ansi);
                        break;

                    //Underlining
                    case 0x2A3E:
                        appendValueElement(parent, "u", lowerFirstChar(((Global.UnderlineCode)sprm.Arguments[0]).ToString()), true);
                        break;

                    //char width
                    case 0x4852:
                        appendValueElement(parent, "w", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //animation
                    case 0x2859:
                        appendValueElement(parent, "effect", ((Global.TextAnimation)sprm.Arguments[0]).ToString(), true);
                        break;

                    default:
                        break;
                }
            }

            //apend lang
            if (lang.Attributes.Count > 0)
            {
                parent.AppendChild(lang);
            }

            //append fonts
            if (rFonts.Attributes.Count > 0)
            {
                parent.AppendChild(rFonts);
            }

            //append color
            if (colorVal.Value != "")
            {
                color.Attributes.Append(colorVal);
                parent.AppendChild(color);
            }
        }

        /// <summary>
        /// CHPX flags are special flags because the can be 0,1,128 and 129,
        /// so this method overrides the appendFlagElement method.
        /// </summary>
        protected override void appendFlagElement(XmlElement node, SinglePropertyModifier sprm, string elementName, bool unique)
        {
            byte flag = sprm.Arguments[0];

            if(flag != 128)
            {
                XmlElement ele = _nodeFactory.CreateElement("w", elementName, OpenXmlNamespaces.WordprocessingML);
                XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);

                if (flag == 0)
                {
                    val.Value = "false";
                    ele.Attributes.Append(val);
                }
                else if (flag == 1)
                {
                    //dont append attribute val
                    //no val attribute means "true"
                }
                else if(flag == 129)
                {
                    //_writer.WriteComment("The flag " + elementName + " had value 129 (style " + _currentIstd + ")");

                    //means that the value is the negation of the style's value
                    if (_currentIstd == UInt16.MaxValue)
                    {
                        //there is NonSerializedAttribute style
                        //supposed the value is false, set it to true
                        //dont append attribute val
                        //no val attribute means "true"
                    }
                    else
                    {
                        StyleSheetDescription std = _doc.Styles.Styles[_currentIstd];
                        foreach (SinglePropertyModifier styleSprm in std.chpx.grpprl)
                        {
                            //find the value in the style
                            if (styleSprm.OpCode == sprm.OpCode)
                            {
                                //negate it
                                byte styleFlag = styleSprm.Arguments[0];
                                switch (styleFlag)
	                            {
                                    case 1:
                                        val.Value = "false";
                                        ele.Attributes.Append(val);
                                        break;
                                    case 0:
                                        //dont append attribute val
                                        //no val attribute means "true"
                                        break;
                                    default:
                                        val.Value = styleFlag.ToString();
                                        ele.Attributes.Append(val);
                                        break;
                                }
                            }
                        }
                    }
                }


                if (unique)
                {
                    foreach (XmlElement exEle in node.ChildNodes)
                    {
                        if (exEle.Name == ele.Name)
                        {
                            node.RemoveChild(exEle);
                            break;
                        }
                    }
                }
                node.AppendChild(ele);
            }
        }

        private string lowerFirstChar(string s)
        {
            return s.Substring(0, 1).ToLower() + s.Substring(1, s.Length - 1);
        }
    }
}
