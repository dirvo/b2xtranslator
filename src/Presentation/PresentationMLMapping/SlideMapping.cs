using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class SlideMapping : PresentationMapping<Slide>
    {
        public SlideMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddSlidePart())
        {
        }

        override public void Apply(Slide slide)
        {
            Console.WriteLine("SlideMapping.Apply");

            RoundTripContentMasterId12 masterInfo = slide.FirstChildWithType<RoundTripContentMasterId12>();
            if (masterInfo != null)
            {
                int mainMasterIdx = (int)masterInfo.MainMasterId - 1;
                int slideLayoutIdx = (int)masterInfo.ContentMasterInstanceId - 1;
                SlideMasterPart master = _ctx.Pptx.PresentationPart.SlideMasterParts[mainMasterIdx];
                SlideLayoutPart layout = master.SlideLayoutParts[slideLayoutIdx];

                this.targetPart.ReferencePart<SlideLayoutPart>(layout);
            }
            else
            {
                // TODO...
            }

            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "sld", OpenXmlNamespaces.PresentationML);
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            // TODO: Write slide data of master slide
            _writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);
            new ShapeTreeMapping(_ctx, _writer).Apply(slide.FirstChildWithType<PPDrawing>());
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            // TODO: Write clrMapOvr

            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
