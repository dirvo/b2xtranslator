using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Reflection;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class MainMasterMapping : PresentationMapping<MainMaster>
    {
        public MainMasterMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddSlideMasterPart())
        {
        }

        override public void Apply(MainMaster master)
        {
            TraceLogger.DebugInternal("MainMasterMapping.Apply");

            SlideMasterPart masterPart = (SlideMasterPart)this.targetPart;

            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "sldMaster", OpenXmlNamespaces.PresentationML);
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            _writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);
            new ShapeTreeMapping(_ctx, _writer).Apply(master.FirstChildWithType<PPDrawing>());
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            // Write clrMap
            ColorMappingAtom clrMap = master.FirstChildWithType<ColorMappingAtom>();
            if (clrMap != null)
            {
                // clrMap from ColorMappingAtom wrongly uses namespace DrawingML
                _writer.WriteStartElement("p", "clrMap", OpenXmlNamespaces.PresentationML);

                foreach (XmlAttribute attr in clrMap.XmlDocumentElement.Attributes)
                    if (attr.Prefix != "xmlns")
                        _writer.WriteAttributeString(attr.LocalName, attr.Value);

                _writer.WriteEndElement();
            }
            else
            {
                // In absence of ColorMappingAtom write default clrMap
                Utils.GetDefaultDocument("clrMap").WriteTo(_writer);
            }


            // Write slide layouts
            _writer.WriteStartElement("p", "sldLayoutIdLst", OpenXmlNamespaces.PresentationML);

            List<RoundTripContentMasterInfo12> slideLayouts = master.AllChildrenWithType<RoundTripContentMasterInfo12>();

            slideLayouts.Sort(delegate(RoundTripContentMasterInfo12 a, RoundTripContentMasterInfo12 b) {
                return a.Instance.CompareTo(b.Instance);
            });

            if (slideLayouts.Count > 0)
            {
                foreach (RoundTripContentMasterInfo12 slideLayout in slideLayouts)
                {
                    SlideLayoutPart layoutPart = masterPart.AddSlideLayoutPart();

                    layoutPart.ReferencePart<SlideMasterPart>(masterPart);

                    slideLayout.XmlDocumentElement.WriteTo(layoutPart.XmlWriter);
                    layoutPart.XmlWriter.Flush();

                    _writer.WriteStartElement("p", "sldLayoutId", OpenXmlNamespaces.PresentationML);
                    _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, layoutPart.RelIdToString);
                    _writer.WriteEndElement();
                }
            }
            else
            {
                // TODO (pre PP2007)
                //throw new NotImplementedException("Write slide Layouts in case of PPT without roundTripContentMasterInfo");
            }

            _writer.WriteEndElement();


            // Write txStyles
            RoundTripOArtTextStyles12 roundTripTxStyles = master.FirstChildWithType<RoundTripOArtTextStyles12>();
            if (roundTripTxStyles != null)
            {
                roundTripTxStyles.XmlDocumentElement.WriteTo(_writer);
            }
            else
            {
                //throw new NotImplementedException("Write txStyles in case of PPT without roundTripTxStyles"); // TODO (pre PP2007)
            }


            // Write theme
            ThemePart themePart = masterPart.AddThemePart();

            Theme theme = master.FirstChildWithType<Theme>();
            if (theme != null)
            {
                theme.XmlDocumentElement.WriteTo(themePart.XmlWriter);
            }
            else
            {
                // In absence of Theme record write default clrMap
                Utils.GetDefaultDocument("theme").WriteTo(themePart.XmlWriter);
            }

            themePart.XmlWriter.Flush();

            PresentationPart presentationPart = _ctx.Pptx.PresentationPart;
            presentationPart.ReferencePart(themePart);

            
            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
