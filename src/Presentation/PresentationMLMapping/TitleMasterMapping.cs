using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class TitleMasterMapping : PresentationMapping<Slide>
    {
        public TitleMasterMapping(ConversionContext ctx, SlideLayoutPart part)
            : base(ctx, part)
        {

        }

        override public void Apply(Slide slide)
        {
            TraceLogger.DebugInternal("TitleMasterMapping.Apply");

            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "sldLayout", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("showMasterSp", "0");
            _writer.WriteAttributeString("type", "title");
            _writer.WriteAttributeString("preserve", "1");

            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            _writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);
            new ShapeTreeMapping(_ctx, _writer).Apply(slide.FirstChildWithType<PPDrawing>());
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
