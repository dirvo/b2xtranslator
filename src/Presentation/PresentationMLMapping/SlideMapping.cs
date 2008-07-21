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
    public class SlideMapping : PresentationMapping<Slide>
    {
        public Slide Slide;

        public SlideMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddSlidePart())
        {
        }

        /// <summary>
        /// Get the id of our real main master.
        /// 
        /// This need not be the id of our immediate master as it can be a title master.
        /// </summary>
        /// <param name="slideAtom">SlideAtom of slide to find main master id for</param>
        /// <returns>Id of main master</returns>
        private UInt32 GetMainMasterId(SlideAtom slideAtom)
        {
            Slide masterSlide = _ctx.Ppt.FindMasterRecordById(slideAtom.MasterId);
            
            // Is our immediate master a title master?
            if (!(masterSlide is MainMaster))
            {
                // Then our main master is the title master's master
                SlideAtom titleSlideAtom = masterSlide.FirstChildWithType<SlideAtom>();
                return titleSlideAtom.MasterId;
            }

            return slideAtom.MasterId;
        }

        override public void Apply(Slide slide)
        {
            this.Slide = slide;
            TraceLogger.DebugInternal("SlideMapping.Apply");

            // Associate slide with slide layout
            SlideAtom slideAtom = slide.FirstChildWithType<SlideAtom>();
            UInt32 mainMasterId = GetMainMasterId(slideAtom);
            MasterLayoutManager layoutManager = _ctx.GetOrCreateLayoutManagerByMasterId(mainMasterId);

            SlideLayoutPart layoutPart;
            RoundTripContentMasterId12 masterInfo = slide.FirstChildWithType<RoundTripContentMasterId12>();

            // PPT2007 OOXML-Layout
            if (masterInfo != null)
            {
                layoutPart = layoutManager.GetLayoutPartByInstanceId(masterInfo.ContentMasterInstanceId);
            }
            // Pre-PPT2007 Title master layout
            else if (mainMasterId != slideAtom.MasterId)
            {
                layoutPart = layoutManager.GetOrCreateLayoutPartForTitleMasterId(slideAtom.MasterId);
            }
            // Pre-PPT2007 SSlideLayoutAtom primitive SlideLayoutType layout
            else
            {
                layoutPart = layoutManager.GetOrCreateLayoutPartByLayoutType((uint)slideAtom.Layout.Geom);
            }

            this.targetPart.ReferencePart(layoutPart);

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
