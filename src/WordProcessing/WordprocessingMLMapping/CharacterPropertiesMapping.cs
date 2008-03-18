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
        private StyleSheet _styleSheet;
        private List<FontFamilyName> _fontTable;
        private XmlElement _rPr;
        private UInt16 _currentIstd;

        public CharacterPropertiesMapping(XmlWriter writer, StyleSheet styles, List<FontFamilyName> fontTable)
            : base(writer)
        {
            _styleSheet = styles;
            _fontTable = fontTable;
            _rPr = _nodeFactory.CreateElement("w", "rPr", OpenXmlNamespaces.WordprocessingML);
        }

        public void Apply(CharacterPropertyExceptions chpx)
        {
            XmlElement shd = _nodeFactory.CreateElement("w", "shd", OpenXmlNamespaces.WordprocessingML);
            XmlElement rFonts = _nodeFactory.CreateElement("w", "rFonts", OpenXmlNamespaces.WordprocessingML);
            XmlElement color = _nodeFactory.CreateElement("w", "color", OpenXmlNamespaces.WordprocessingML);
            XmlAttribute colorVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);


            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //style id 
                    case 0x4A30:
                        _currentIstd = System.BitConverter.ToUInt16(sprm.Arguments, 0);
                        appendValueElement(_rPr, "rStyle", StyleSheetMapping.MakeStyleId(_styleSheet.Styles[_currentIstd].xstzName));
                        break;

                    //Element flags
                    case 0x085A:
                        appendFlagElement(sprm, "rtl");
                        break;
                    case 0x0835:
                        appendFlagElement(sprm, "b");
                        break;
                    case 0x085C:
                        appendFlagElement(sprm, "bCs");
                        break;
                    case 0x083B:
                        appendFlagElement(sprm, "caps");
                        break;
                    case 0x0882:
                        appendFlagElement(sprm, "cs");
                        break;
                    case 0x2A53:
                        appendFlagElement(sprm, "dstrike");
                        break;
                    case 0x0858:
                        appendFlagElement(sprm, "emboss");
                        break;
                    case 0x0854:
                        appendFlagElement(sprm, "imprint");
                        break;
                    case 0x0836:
                        appendFlagElement(sprm, "i");
                        break;
                    case 0x085D:
                        appendFlagElement(sprm, "iCs");
                        break;
                    case 0x0875:
                        appendFlagElement(sprm, "noProof");
                        break;
                    case 0x0838:
                        appendFlagElement(sprm, "outline");
                        break;
                    case 0x0839:
                        appendFlagElement(sprm, "shadow");
                        break;
                    case 0x083A:
                        appendFlagElement(sprm, "smallCaps");
                        break;
                    case 0x0818:
                        appendFlagElement(sprm, "specVanish");
                        break;
                    case 0x0837:
                        appendFlagElement(sprm, "strike");
                        break;
                    case 0x083C:
                        appendFlagElement(sprm, "vanish");
                        break;
                    case 0x0811:
                        appendFlagElement(sprm, "webHidden");
                        break;
                    
                    //color
                    case 0x2A42:
                    case 0x4A60:
                        colorVal.Value = ((Global.ColorIdentifier)(sprm.Arguments[0])).ToString();
                        break;
                    case 0x6870:
                        colorVal.Value = String.Format("{0:x6}", System.BitConverter.ToUInt32(sprm.Arguments, 0));
                        break;

                    //highlightning
                    case 0x2A0C:
                        appendValueElement(_rPr, "highlight", ((Global.ColorIdentifier)sprm.Arguments[0]).ToString());
                        break;

                    //spacing
                    case 0x8840:
                        appendValueElement(_rPr, "spacing", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //font size
                    case 0x4A43:
                        appendValueElement(_rPr, "sz", sprm.Arguments[0].ToString());
                        break;
                    case 0x484B:
                        appendValueElement(_rPr, "kern", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x4A61:
                        appendValueElement(_rPr, "szCs", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //font family
                    case 0x4A4F:
                       XmlAttribute ascii = _nodeFactory.CreateAttribute("w", "ascii", OpenXmlNamespaces.WordprocessingML);
                       ascii.Value = _fontTable[System.BitConverter.ToUInt16(sprm.Arguments, 0)].xszFtn;
                       rFonts.Attributes.Append(ascii);
                       break;
                   case 0x4A50:
                       XmlAttribute eastAsia = _nodeFactory.CreateAttribute("w", "eastAsia", OpenXmlNamespaces.WordprocessingML);
                       eastAsia.Value = _fontTable[System.BitConverter.ToUInt16(sprm.Arguments, 0)].xszFtn;
                       rFonts.Attributes.Append(eastAsia);
                       break;
                    case 0x4A51:
                        XmlAttribute ansi = _nodeFactory.CreateAttribute("w", "hAnsi", OpenXmlNamespaces.WordprocessingML);
                        ansi.Value = _fontTable[System.BitConverter.ToUInt16(sprm.Arguments, 0)].xszFtn;
                        rFonts.Attributes.Append(ansi);
                        break;

                    //Underlining
                    case 0x2A3E:
                        appendValueElement(_rPr, "u", lowerFirstChar(((Global.UnderlineCode)sprm.Arguments[0]).ToString()));
                        break;

                    //char width
                    case 0x4852:
                        appendValueElement(_rPr, "w", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //animation
                    case 0x2859:
                        appendValueElement(_rPr, "effect", ((Global.TextAnimation)sprm.Arguments[0]).ToString());
                        break;

                    default:
                        break;
                }
            }

            //append fonts
            if (rFonts.Attributes.Count > 0)
            {
                _rPr.AppendChild(rFonts);
            }

            //append color
            if (colorVal.Value != "")
            {
                color.Attributes.Append(colorVal);
                _rPr.AppendChild(color);
            }
            
            //write properties
            if (_rPr.ChildNodes.Count > 0 || _rPr.Attributes.Count > 0)
            {
                _rPr.WriteTo(_writer);
            }
        }

        private string lowerFirstChar(string s)
        {
            return s.Substring(0, 1).ToLower() + s.Substring(1, s.Length - 1);
        }

        private void appendFlagElement(SinglePropertyModifier sprm, string elementName)
        {
            byte b = sprm.Arguments[0];

            //value 
            if(b == 129)
            {
                byte bStyle = 0;

                //find style's value
                foreach (SinglePropertyModifier sprmStyle in _styleSheet.Styles[_currentIstd].chpx.grpprl)
                {
                    if(sprm.OpCode == sprmStyle.OpCode)
                    {
                        bStyle = sprmStyle.Arguments[0];
                    }
                }

                //value is set to the negation of the style's value
                switch (bStyle)
                {
                    case 0:
                        b = 1;
                        break;
                    case 1:
                        b = 0;
                        break;
                    default:
                        b = bStyle;
                        break;
                }
            }
            
            if(b == 128)
            {
                //value is set to the value of the style
                //don't set anything, value will be inherited from style
            }
            else
            {
                XmlElement ele = _nodeFactory.CreateElement("w", elementName, OpenXmlNamespaces.WordprocessingML);
                XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                val.Value = b.ToString();
                ele.Attributes.Append(val);
                _rPr.AppendChild(ele);
            }
        }
    }
}
