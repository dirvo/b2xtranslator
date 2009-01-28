/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

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
                layoutPart = layoutManager.GetOrCreateLayoutPartByLayoutType(slideAtom.Layout.Geom, slideAtom.Layout.PlaceholderTypes);
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

            //TODO: write background properties (p:bg)

            _writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);

            ShapeTreeMapping stm = new ShapeTreeMapping(_ctx, _writer);
            stm.parentSlideMapping = this;
            stm.Apply(slide.FirstChildWithType<PPDrawing>());
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
