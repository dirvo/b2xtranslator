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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class PresentationPartMapping : PresentationMapping<PowerpointDocument>
    {
        public List<SlideMapping> SlideMappings = new List<SlideMapping>();
        private List<NoteMapping> NoteMappings = new List<NoteMapping>();

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
            CreateNotesMasters(ppt);
            CreateHandoutMasters(ppt);
            CreateSlides(ppt, documentRecord);
            //CreatePictures(ppt, documentRecord);
                        
            WriteMainMasters(ppt);
            WriteSlides(ppt, documentRecord);

            // sldSz and notesSz
            WriteSizeInfo(ppt, documentRecord);

            WriteDefaultTextStyle(ppt, documentRecord);

            // End the document
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        private void WriteDefaultTextStyle(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            _writer.WriteStartElement("p", "defaultTextStyle", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("a", "defPPr", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "defRPr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("lang", "en-US");
            _writer.WriteEndElement(); //defRPr
            _writer.WriteEndElement(); //defPPr


            TextMasterStyleAtom defaultStyle = _ctx.Ppt.DocumentRecord.FirstChildWithType<DIaLOGIKa.b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<TextMasterStyleAtom>();

            TextMasterStyleMapping map = new TextMasterStyleMapping(_ctx, _writer, null);
            
            for (int i = 0; i < defaultStyle.IndentLevelCount; i++)
            {
                map.writepPr(defaultStyle.CRuns[i], defaultStyle.PRuns[i], null, i, false, true);
            }
            for (int i = defaultStyle.IndentLevelCount; i < 9; i++)
            {
                map.writepPr(defaultStyle.CRuns[0], defaultStyle.PRuns[0], null, i, false, true);
            }


            _writer.WriteEndElement(); //defaultTextStyle
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

        //private void CreatePictures(PowerpointDocument ppt, DocumentContainer documentRecord)
        //{
        //    foreach (Record rec in ppt.PicturesContainer._pictures.Values)
        //    {
        //        switch (rec.TypeCode)
        //        {

        //            case 0xF01D:
        //            case 0xF01E:
        //            case 0xF01F:
        //            case 0xF020:
        //            case 0xF021:
        //                BitmapBlip b = (BitmapBlip)rec;
        //                ImagePart imgPart = null;
        //                imgPart = this.targetPart.AddImagePart(ImagePart.ImageType.Png);
        //                System.IO.Stream outStream = imgPart.GetStream();
        //                outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);

        //                break;

        //            default:
        //                break;
        //        }
        //    }
        //}

        
       private void CreateSlides(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            foreach (SlideListWithText lst in ppt.DocumentRecord.AllChildrenWithType<SlideListWithText>())
            {
                if (lst.Instance == 0)
                {
                    foreach (SlidePersistAtom at in lst.SlidePersistAtoms)
                    {
                        foreach (Slide slide in ppt.SlideRecords)
                        {
                            if (slide.PersistAtom == at)
                            {
                                SlideMapping sMapping = new SlideMapping(_ctx);
                                sMapping.Apply(slide);
                                this.SlideMappings.Add(sMapping);
                            }
                        }


                    }
                }
                bool found = false;
                if (lst.Instance == 2) //notes
                {
                    foreach (SlidePersistAtom at in lst.SlidePersistAtoms)
                    {
                        found = false;
                        foreach (Note note in ppt.NoteRecords)
                        {
                            if (note.PersistAtom.SlideId == at.SlideId)
                            {
                                NotesAtom a = note.FirstChildWithType<NotesAtom>();
                                foreach (SlideMapping slideMapping in this.SlideMappings)
                                {
                                    if (slideMapping.Slide.PersistAtom.SlideId == a.SlideIdRef)
                                    {
                                        NoteMapping nMapping = new NoteMapping(_ctx, slideMapping);
                                        nMapping.Apply(note);
                                        this.NoteMappings.Add(nMapping);
                                        found = true;
                                    }
                                }
                                
                            }
                        }
                        if (!found)
                        {
                            string s = "";
                        }

                    }
                }
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

        private void CreateNotesMasters(PowerpointDocument ppt)
        {
            foreach (Note m in ppt.NotesMasterRecords)
            {
                _ctx.GetOrCreateNotesMasterMappingByMasterId(0).Apply(m);
            }
        }

        private void CreateHandoutMasters(PowerpointDocument ppt)
        {
            foreach (Handout m in ppt.HandoutMasterRecords)
            {
                _ctx.GetOrCreateHandoutMasterMappingByMasterId(0).Apply(m);
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

            WriteNoteMaster(ppt);
            WriteHandoutMaster(ppt);
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

        private void WriteNoteMaster(PowerpointDocument ppt)
        {
            if (ppt.NotesMasterRecords.Count > 0)
            {
                _writer.WriteStartElement("p", "notesMasterIdLst", OpenXmlNamespaces.PresentationML);

                foreach (Note m in ppt.NotesMasterRecords)
                {
                    this.WriteNoteMaster2(ppt, m);
                }

                _writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Writes a notes master.
        /// 
        /// <param name="ppt">PowerpointDocument record</param>
        /// <param name="m">Notes master record</param>
        private void WriteNoteMaster2(PowerpointDocument ppt, Note m)
        {
            _writer.WriteStartElement("p", "notesMasterId", OpenXmlNamespaces.PresentationML);

            NotesMasterMapping mapping = _ctx.GetOrCreateNotesMasterMappingByMasterId(0);
            mapping.Write();

            string relString = mapping.targetPart.RelIdToString;

            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            _writer.WriteEndElement();
        }

        private void WriteHandoutMaster(PowerpointDocument ppt)
        {
            if (ppt.HandoutMasterRecords.Count > 0)
            {

                _writer.WriteStartElement("p", "handoutMasterIdLst", OpenXmlNamespaces.PresentationML);

                foreach (Handout m in ppt.HandoutMasterRecords)
                {
                    this.WriteHandoutMaster2(ppt, m);
                }

                _writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Writes a handout master.
        /// 
        /// <param name="ppt">PowerpointDocument record</param>
        /// <param name="m">Handout master record</param>
        private void WriteHandoutMaster2(PowerpointDocument ppt, Handout m)
        {
            _writer.WriteStartElement("p", "handoutMasterId", OpenXmlNamespaces.PresentationML);

            HandoutMasterMapping mapping = _ctx.GetOrCreateHandoutMasterMappingByMasterId(0);
            mapping.Write();

            string relString = mapping.targetPart.RelIdToString;

            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            _writer.WriteEndElement();
        }
    }
}
