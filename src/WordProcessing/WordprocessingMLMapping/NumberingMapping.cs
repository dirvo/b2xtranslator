using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class NumberingMapping : AbstractOpenXmlMapping,
          IMapping<ListTable>
    {

        public NumberingMapping(NumberingDefinitionsPart numPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(numPart.GetStream(), xws))
        {
        }

        public void Apply(ListTable rglst)
        {
            _writer.WriteStartElement("w", "numbering", OpenXmlNamespaces.WordprocessingML);

            for (int i = 0 ; i < rglst.Count; i++)
            {
                ListData lstf = rglst[i];

                //start abstractNum
                _writer.WriteStartElement("w", "abstractNum", OpenXmlNamespaces.WordprocessingML);

                _writer.WriteAttributeString("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML, i.ToString());

                //nsid
                _writer.WriteStartElement("w", "nsid", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, String.Format("{0:X8}", lstf.lsid));
                _writer.WriteEndElement();

                //template
                _writer.WriteStartElement("w", "tmpl", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, String.Format("{0:X8}", lstf.tplc));
                _writer.WriteEndElement();

                //multiLevelType
                _writer.WriteStartElement("w", "multiLevelType", OpenXmlNamespaces.WordprocessingML);
                if (lstf.fHybrid)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "hybridMultilevel");
                else if (lstf.fSimpleList)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "singleLevel");
                else
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "multiLevel");
                _writer.WriteEndElement();

                //end abstractNum
                _writer.WriteEndElement();

                //start num
                _writer.WriteStartElement("w", "num", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "numId", OpenXmlNamespaces.WordprocessingML, (i+1).ToString());
                _writer.WriteStartElement("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, i.ToString());
                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
            _writer.Flush();
        }
    }
}
