using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class CharacterPropertiesMapping : AbstractOpenXmlMapping,
          IMapping<CharacterPropertyExceptions>
    {
        private StyleSheet _styleSheet;
        private XmlNode _rPr;

        public CharacterPropertiesMapping(XmlWriter writer, StyleSheet styles)
            : base(writer)
        {
            _styleSheet = styles;
            _rPr = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "rPr", OpenXmlNamespaces.WordprocessingML);
        }

        public void Apply(CharacterPropertyExceptions chpx)
        {
            //write properties
            if (_rPr.ChildNodes.Count > 0 || _rPr.Attributes.Count > 0)
            {
                _rPr.WriteTo(_writer);
            }
        }
    }
}
