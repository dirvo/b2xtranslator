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
    public class MasterMapping : PresentationMapping<Slide>
    {
        public SlideMasterPart MasterPart;
        protected Slide Master;
        protected UInt32 MasterId;
        protected MasterLayoutManager LayoutManager;

        public MasterMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddSlideMasterPart())
        {
            this.MasterPart = (SlideMasterPart)this.targetPart;
        }

        override public void Apply(Slide master)
        {
            TraceLogger.DebugInternal("MasterMapping.Apply");
            UInt32 masterId = master.PersistAtom.SlideId;
            _ctx.RegisterMasterMapping(masterId, this);

            this.Master = master;
            this.MasterId = master.PersistAtom.SlideId;
            this.LayoutManager = _ctx.GetOrCreateLayoutManagerByMasterId(this.MasterId);

            // Add PPT2007 roundtrip slide layouts
            List<RoundTripContentMasterInfo12> rtSlideLayouts = this.Master.AllChildrenWithType<RoundTripContentMasterInfo12>();

            foreach (RoundTripContentMasterInfo12 slideLayout in rtSlideLayouts)
            {
                SlideLayoutPart layoutPart = this.LayoutManager.AddLayoutPartWithInstanceId(slideLayout.Instance);

                slideLayout.XmlDocumentElement.WriteTo(layoutPart.XmlWriter);
                layoutPart.XmlWriter.Flush();
            }
        }

        public void Write()
        {
            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "sldMaster", OpenXmlNamespaces.PresentationML);
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            _writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);
            new ShapeTreeMapping(_ctx, _writer).Apply(this.Master.FirstChildWithType<PPDrawing>());
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            // Write clrMap
            ColorMappingAtom clrMap = this.Master.FirstChildWithType<ColorMappingAtom>();
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

            // Write slide layout part id list
            _writer.WriteStartElement("p", "sldLayoutIdLst", OpenXmlNamespaces.PresentationML);

            List<SlideLayoutPart> layoutParts = this.LayoutManager.GetAllLayoutParts();

            // Maser must have at least one SlideLayout or RepairDialog will appear
            if (layoutParts.Count == 0)
            {
                SlideLayoutPart layoutPart = this.LayoutManager.GetOrCreateLayoutPartByLayoutType(0, null);
                layoutParts.Add(layoutPart);
            }

            foreach (SlideLayoutPart slideLayoutPart in layoutParts)
            {
                _writer.WriteStartElement("p", "sldLayoutId", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, slideLayoutPart.RelIdToString);
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();

            // Write txStyles
            RoundTripOArtTextStyles12 roundTripTxStyles = this.Master.FirstChildWithType<RoundTripOArtTextStyles12>();
            if (roundTripTxStyles != null)
            {
                roundTripTxStyles.XmlDocumentElement.WriteTo(_writer);
            }
            else
            {
                //throw new NotImplementedException("Write txStyles in case of PPT without roundTripTxStyles"); // TODO (pre PP2007)
            }

            // Write theme
            //
            // Note: We need to create a new theme part for each master,
            // even if it they have the same content.
            //
            // Otherwise PPT will complain about the structure of the file.
            ThemePart themePart = _ctx.Pptx.PresentationPart.AddThemePart();

            XmlNode xmlDoc;
            Theme theme = this.Master.FirstChildWithType<Theme>();

            if (theme != null)
            {
                xmlDoc = theme.XmlDocumentElement;
            }
            else
            {
                // In absence of Theme record use default theme
                xmlDoc = Utils.GetDefaultDocument("theme");
            }

            xmlDoc.WriteTo(themePart.XmlWriter);
            themePart.XmlWriter.Flush();

            this.MasterPart.ReferencePart(themePart);
            
            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
