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

        public TableRowPropertiesMapping(XmlWriter writer)
            : base(writer)
        {
            _trPr = _nodeFactory.CreateElement("w", "trPr", OpenXmlNamespaces.WordprocessingML);
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            foreach (SinglePropertyModifier sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //row height
                    case 0x9407:
                        XmlElement rowHeight = _nodeFactory.CreateElement("w", "trHeight", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute rowHeightVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        rowHeightVal.Value = System.BitConverter.ToInt16(sprm.Arguments, 0).ToString();
                        rowHeight.Attributes.Append(rowHeightVal);
                        XmlAttribute rowHeightRule = _nodeFactory.CreateAttribute("w", "hRule", OpenXmlNamespaces.WordprocessingML);
                        rowHeightRule.Value = "exact";
                        rowHeight.Attributes.Append(rowHeightRule);
                        _trPr.AppendChild(rowHeight);
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
