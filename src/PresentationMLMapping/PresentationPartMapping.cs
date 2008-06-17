using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class PresentationPartMapping : PresentationMapping<PowerpointDocument>
    {
        public PresentationPartMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart)
        {
        }

        public override void Apply(PowerpointDocument ppt)
        {
            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "presentation", OpenXmlNamespaces.PresentationML);

            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships); 

            WriteSlideMasters(ppt);
            WriteSlides(ppt);

            // sldSz and notesSz
            WriteSizeInfo(ppt);

            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        private void WriteSizeInfo(PowerpointDocument ppt)
        {
            DocumentAtom doc = ppt.FirstRootRecordWithType<PptDocumentRecord>().FirstChildWithType<DocumentAtom>();

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

            _writer.WriteAttributeString("p", "cx", OpenXmlNamespaces.PresentationML, notesWidth.ToString());
            _writer.WriteAttributeString("p", "cy", OpenXmlNamespaces.PresentationML, notesHeight.ToString());

            _writer.WriteEndElement();
        }

        private void WriteSlideSizeInfo(DocumentAtom doc)
        {
            int slideWidth = Utils.MasterCoordToEMU(doc.SlideSize.X);
            int slideHeight = Utils.MasterCoordToEMU(doc.SlideSize.Y);
            string slideType = Utils.SlideSizeTypeToXMLValue(doc.SlideSizeType);

            _writer.WriteStartElement("p", "sldSz", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("p", "cx", OpenXmlNamespaces.PresentationML, slideWidth.ToString());
            _writer.WriteAttributeString("p", "cy", OpenXmlNamespaces.PresentationML, slideHeight.ToString());
            _writer.WriteAttributeString("p", "type", OpenXmlNamespaces.PresentationML, slideType);

            _writer.WriteEndElement();

        }

        private void WriteSlides(PowerpointDocument ppt)
        {
            _writer.WriteStartElement("p", "sldIdLst", OpenXmlNamespaces.PresentationML);

            foreach (Slide slide in ppt.AllRootRecordsWithType<Slide>())
            {
                WriteSlide(slide);
            }

            _writer.WriteEndElement();
        }

        private void WriteSlide(Slide slide)
        {
            _writer.WriteStartElement("p", "sldId", OpenXmlNamespaces.PresentationML);

            SlideMapping mapping = new SlideMapping(_ctx);
            mapping.Apply(slide);

            string relString = mapping.targetPart.RelIdToString;

            // FIXME: Get the real id from the presentation tree? (Where is it stored?)
            _writer.WriteAttributeString("p", "id", OpenXmlNamespaces.PresentationML, relString);
            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            _writer.WriteEndElement();
        }

        private void WriteSlideMasters(PowerpointDocument ppt)
        {
            _writer.WriteStartElement("p", "sldMasterIdLst", OpenXmlNamespaces.PresentationML);

            foreach (MainMaster mm in ppt.AllRootRecordsWithType<MainMaster>())
            {
                WriteSlideMaster(mm);
            }

            _writer.WriteEndElement();
        }

        private void WriteSlideMaster(MainMaster mm)
        {
            _writer.WriteStartElement("p", "sldMasterId", OpenXmlNamespaces.PresentationML);

            MainMasterMapping mapping = new MainMasterMapping(_ctx);
            mapping.Apply(mm);

            string relString = mapping.targetPart.RelIdToString;

            // FIXME: Get the real id from the presentation tree? (Where is it stored?)
            _writer.WriteAttributeString("p", "id", OpenXmlNamespaces.PresentationML, relString);
            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            _writer.WriteEndElement();
        }
    }
}
