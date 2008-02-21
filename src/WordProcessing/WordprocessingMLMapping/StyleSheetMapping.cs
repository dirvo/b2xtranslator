using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Diagnostics;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class StyleSheetMapping 
        : AbstractOpenXmlMapping,
          IMapping<StyleSheet>,
          IMapping<StyleSheetDescription>
    {
        public StyleSheetMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(StyleSheet visited)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "styles", OpenXmlNamespaces.WordprocessingML);


            foreach (StyleSheetDescription style in visited.Styles)
            {
                if (style != null)
                {
                    style.Convert(this);
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndDocument();
        }

        
        public void Apply(StyleSheetDescription visited)
        {
            _writer.WriteStartElement("w", "style", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "styleId", visited.xstzName);

            _writer.WriteEndElement();
        }
    }
}
