using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    class TableRowPropertiesMapping :
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private XmlElement _trPr;
        private CharacterPropertyExceptions _rowEndChpx;

        public TableRowPropertiesMapping(XmlWriter writer, CharacterPropertyExceptions rowEndChpx)
            : base(writer)
        {
            _trPr = _nodeFactory.CreateElement("w", "trPr", OpenXmlNamespaces.WordprocessingML);
            _rowEndChpx = rowEndChpx;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            //delete infos
            if (_rowEndChpx != null && _rowEndChpx.IsDeleted)
            {
                XmlElement del = _nodeFactory.CreateElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                _trPr.AppendChild(del);
            }

            foreach (SinglePropertyModifier sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)  
                {
                    //width after
                    case 0xF618:
                        XmlElement wAfter = _nodeFactory.CreateElement("w", "wAfter", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute wAfterValue = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                        wAfterValue.Value = System.BitConverter.ToInt16(sprm.Arguments, 1).ToString();
                        wAfter.Attributes.Append(wAfterValue);
                        XmlAttribute wAfterType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                        wAfterType.Value = "dxa";
                        wAfter.Attributes.Append(wAfterType);
                        _trPr.AppendChild(wAfter);
                        break;

                    //width before
                    case 0xF617:
                        Int16 before = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        if (before != 0)
                        {
                            XmlElement wBefore = _nodeFactory.CreateElement("w", "wBefore", OpenXmlNamespaces.WordprocessingML);
                            XmlAttribute wBeforeValue = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                            wBeforeValue.Value = before.ToString();
                            wBefore.Attributes.Append(wBeforeValue);
                            XmlAttribute wBeforeType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                            wBeforeType.Value = "dxa";
                            wBefore.Attributes.Append(wBeforeType);
                            _trPr.AppendChild(wBefore);
                        }
                        break;

                    case 0x9602:
                        break;

                    //row height
                    case 0x9407:
                        XmlElement rowHeight = _nodeFactory.CreateElement("w", "trHeight", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute rowHeightVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute rowHeightRule = _nodeFactory.CreateAttribute("w", "hRule", OpenXmlNamespaces.WordprocessingML);
                        Int16 rH = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        if (rH > 0)
                        {
                            rowHeightRule.Value = "atLeast";
                        }
                        else
                        {
                            rowHeightRule.Value = "exact";
                            rH *= -1;
                        }
                        rowHeightVal.Value = rH.ToString();
                        rowHeight.Attributes.Append(rowHeightVal);
                        rowHeight.Attributes.Append(rowHeightRule);
                        _trPr.AppendChild(rowHeight);
                        break;

                    //can't split
                    case 0x3403:
                        appendFlagElement(_trPr, sprm, "cantSplit", true);
                        break;

                    //conditional formatting
                    case 0xD66A:
                        ;
                        break;

                    //div id
                    case 0x7469:
                        appendValueElement(_trPr, "divId", System.BitConverter.ToInt32(sprm.Arguments, 0).ToString(), true);
                        break;
                }
            }

            //write Properties
            if (_trPr.ChildNodes.Count > 0 || _trPr.Attributes.Count > 0)
            {
                _trPr.WriteTo(_writer);
            }
        }
    }
}
