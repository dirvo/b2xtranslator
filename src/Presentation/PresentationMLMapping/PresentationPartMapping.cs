using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class PresentationPartMapping : PresentationMapping<PowerpointDocument>
    {
        private List<SlideMapping> SlideMappings = new List<SlideMapping>();

        public PresentationPartMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart)
        {
        }

        public override void Apply(PowerpointDocument ppt)
        {
            DocumentContainer documentRecord = ppt.DocumentRecord;

            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "presentation", OpenXmlNamespaces.PresentationML);

            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            CreateMainMasters(ppt);
            CreateSlides(ppt, documentRecord);

            WriteMainMasters(ppt);
            WriteSlides(ppt, documentRecord);

            // sldSz and notesSz
            WriteSizeInfo(ppt, documentRecord);

            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        private void WriteSizeInfo(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            DocumentAtom doc = documentRecord.FirstChildWithType<DocumentAtom>();

            // Write slide size and type
            WriteSlideSizeInfo(doc);

            // Write notes size
            WriteNotesSizeInfo(doc);
        }

        private void WriteNotesSizeInfo(DocumentAtom doc)
        {
            int notesWidth = Utils.MasterCoordToEMU(doc.NotesSize.X);
            int notesHeight = Utils.MasterCoordToEMU(doc.NotesSize.Y);

            _writer.WriteStartElement("p", "notesSz", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("cx", notesWidth.ToString());
            _writer.WriteAttributeString("cy", notesHeight.ToString());

            _writer.WriteEndElement();
        }

        private void WriteSlideSizeInfo(DocumentAtom doc)
        {
            int slideWidth = Utils.MasterCoordToEMU(doc.SlideSize.X);
            int slideHeight = Utils.MasterCoordToEMU(doc.SlideSize.Y);
            string slideType = Utils.SlideSizeTypeToXMLValue(doc.SlideSizeType);

            _writer.WriteStartElement("p", "sldSz", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("cx", slideWidth.ToString());
            _writer.WriteAttributeString("cy", slideHeight.ToString());
            _writer.WriteAttributeString("type", slideType);

            _writer.WriteEndElement();

        }

        private void CreateSlides(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            foreach (Slide slide in ppt.SlideRecords)
            {
                SlideMapping sMapping = new SlideMapping(_ctx);
                sMapping.Apply(slide);
                this.SlideMappings.Add(sMapping);
            }
        }

        private void WriteSlides(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            _writer.WriteStartElement("p", "sldIdLst", OpenXmlNamespaces.PresentationML);

            foreach (SlideMapping sMapping in this.SlideMappings)
            {
                WriteSlide(sMapping);
            }

            _writer.WriteEndElement();
        }

        private void WriteSlide(SlideMapping sMapping)
        {
            Slide slide = sMapping.Slide;

            _writer.WriteStartElement("p", "sldId", OpenXmlNamespaces.PresentationML);

            SlideAtom slideAtom = slide.FirstChildWithType<SlideAtom>();

            _writer.WriteAttributeString("id", slide.PersistAtom.SlideId.ToString());
            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, sMapping.targetPart.RelIdToString);

            _writer.WriteEndElement();
        }

        private void CreateMainMasters(PowerpointDocument ppt)
        {
            foreach (Slide m in ppt.MainMasterRecords)
            {
                _ctx.GetOrCreateMasterMappingByMasterId(m.PersistAtom.SlideId).Apply(m);
            }
        }

        private void WriteMainMasters(PowerpointDocument ppt)
        {
            _writer.WriteStartElement("p", "sldMasterIdLst", OpenXmlNamespaces.PresentationML);

            foreach (MainMaster m in ppt.MainMasterRecords)
            {
                this.WriteMainMaster(ppt, m);
            }

            _writer.WriteEndElement();
        }

        /// <summary>
        /// Writes a slide master.
        /// 
        /// A slide master can either be a main master (type MainMaster) or title master (type Slide).
        /// <param name="ppt">PowerpointDocument record</param>
        /// <param name="m">Main master record</param>
        private void WriteMainMaster(PowerpointDocument ppt, MainMaster m)
        {
            _writer.WriteStartElement("p", "sldMasterId", OpenXmlNamespaces.PresentationML);

            MasterMapping mapping = _ctx.GetOrCreateMasterMappingByMasterId(m.PersistAtom.SlideId);
            mapping.Write();

            string relString = mapping.targetPart.RelIdToString;

            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            _writer.WriteEndElement();
        }
    }
}
