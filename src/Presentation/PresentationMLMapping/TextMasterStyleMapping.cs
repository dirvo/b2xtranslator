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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class TextMasterStyleMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;
        Slide _Master;
        public PresentationMapping<Slide> _parentSlideMapping = null;

        public TextMasterStyleMapping(ConversionContext ctx, XmlWriter writer, PresentationMapping<Slide> parentSlideMapping)
            : base(writer)
        {
            _ctx = ctx;
            _parentSlideMapping = parentSlideMapping;
        }

        public void Apply(Slide Master)
        {

            _Master = Master;

            List<TextMasterStyleAtom> docinfoatoms = _ctx.Ppt.DocumentRecord.FirstChildWithType<DIaLOGIKa.b2xtranslator.PptFileFormat.Environment>().AllChildrenWithType<TextMasterStyleAtom>();
            List<TextMasterStyleAtom> atoms = Master.AllChildrenWithType<TextMasterStyleAtom>();

            List<TextMasterStyleAtom> titleAtoms = new List<TextMasterStyleAtom>();

            List<TextMasterStyle9Atom> body9atoms = new List<TextMasterStyle9Atom>();
            List<TextMasterStyle9Atom> title9atoms = new List<TextMasterStyle9Atom>();
            foreach (ProgTags progtags in Master.AllChildrenWithType<ProgTags>())
	        {
        		foreach (ProgBinaryTag progbinarytag in progtags.AllChildrenWithType<ProgBinaryTag>())
	            {
                    foreach (ProgBinaryTagDataBlob blob in progbinarytag.AllChildrenWithType<ProgBinaryTagDataBlob>())
                    {
                        foreach (TextMasterStyle9Atom atom in blob.AllChildrenWithType<TextMasterStyle9Atom>())
                        {
                            if (atom.Instance == 0) title9atoms.Add(atom);
                            if (atom.Instance == 1) body9atoms.Add(atom);
                        }
                    }            		
	            }
	        }

            List<TextMasterStyleAtom> bodyAtoms = new List<TextMasterStyleAtom>();
            foreach (TextMasterStyleAtom atom in atoms)
            {
                if (atom.Instance == 0) titleAtoms.Add(atom);   
                if (atom.Instance == 1) bodyAtoms.Add(atom);
            }
            
            _writer.WriteStartElement("p", "txStyles", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "titleStyle", OpenXmlNamespaces.PresentationML);

            ParagraphRun9 pr9 = null;
            foreach (TextMasterStyleAtom atom in titleAtoms)
            {
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    if (title9atoms.Count > 0 && title9atoms[0].pruns.Count > i) pr9 = title9atoms[0].pruns[i];
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    if (title9atoms.Count > 0 && title9atoms[0].pruns.Count > i) pr9 = title9atoms[0].pruns[i];
                    writepPr(atom.CRuns[0], atom.PRuns[0], pr9, i);
                }
            }

            _writer.WriteEndElement(); //titleStyle

            _writer.WriteStartElement("p", "bodyStyle", OpenXmlNamespaces.PresentationML);

            foreach (TextMasterStyleAtom atom in bodyAtoms)
            {
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    if (body9atoms.Count > 0 && body9atoms[0].pruns.Count > i) pr9 = body9atoms[0].pruns[i];
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    if (body9atoms.Count > 0 && body9atoms[0].pruns.Count > i) pr9 = body9atoms[0].pruns[i];
                    writepPr(atom.CRuns[0], atom.PRuns[0],pr9, i);
                }
            }

            _writer.WriteEndElement(); //bodyStyle

            _writer.WriteEndElement(); //txStyles
        }

        private void writepPr(CharacterRun cr, ParagraphRun pr, ParagraphRun9 pr9, int IndentLevel)
        {
            _writer.WriteStartElement("a", "lvl" + (IndentLevel+1).ToString() + "pPr", OpenXmlNamespaces.DrawingML);

            if (pr.AlignmentPresent)
            {
                switch (pr.Alignment)
                {
                    case 0x0000: //Left
                        _writer.WriteAttributeString("algn", "l");
                        break;
                    case 0x0001: //Center
                        _writer.WriteAttributeString("algn", "ctr");
                        break;
                    case 0x0002: //Right
                        _writer.WriteAttributeString("algn", "r");
                        break;
                    case 0x0003: //Justify
                        _writer.WriteAttributeString("algn", "just");
                        break;
                    case 0x0004: //Distributed
                        _writer.WriteAttributeString("algn", "dist");
                        break;
                    case 0x0005: //ThaiDistributed
                        _writer.WriteAttributeString("algn", "thaiDist");
                        break;
                    case 0x0006: //JustifyLow
                        _writer.WriteAttributeString("algn", "justLow");
                        break;
                }
            }

            if (pr.TextDirectionPresent)
            {
                switch (pr.TextDirection)
                {
                    case 0x0000:
                        _writer.WriteAttributeString("rtl", "0");
                        break;
                    case 0x0001:
                        _writer.WriteAttributeString("rtl", "1");
                        break;
                }
            }
            else
            {
                _writer.WriteAttributeString("rtl", "0");
            }

            if (pr.FontAlignPresent) 
            {
                switch (pr.FontAlign)
                {
                    case 0x0000: //Roman
                        _writer.WriteAttributeString("fontAlgn", "base");
                        break;
                    case 0x0001: //Hanging
                        _writer.WriteAttributeString("fontAlgn", "t");
                        break;
                    case 0x0002: //Center
                        _writer.WriteAttributeString("fontAlgn", "ctr");
                        break;
                    case 0x0003: //UpholdFixed
                        _writer.WriteAttributeString("fontAlgn", "b");
                        break;
                }                
            }

            if (pr.SpaceBeforePresent)
            {
                _writer.WriteStartElement("a", "spcBef", OpenXmlNamespaces.DrawingML);               
                if (pr.SpaceBefore < 0)
                {
                    _writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", (-1 * 1000 * pr.SpaceBefore).ToString()); //TODO: this has to be verified!
                    _writer.WriteEndElement(); //spcPct
                }
                else
                {
                    _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", (1000 * pr.SpaceBefore).ToString());
                    _writer.WriteEndElement(); //spcPct
                }
                _writer.WriteEndElement(); //spcBef
            }

            if (pr.SpaceAfterPresent)
            {
                _writer.WriteStartElement("a", "spcAft", OpenXmlNamespaces.DrawingML);
                if (pr.SpaceAfter < 0)
                {
                    _writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", (-1 * pr.SpaceAfter).ToString()); //TODO: this has to be verified!
                    _writer.WriteEndElement(); //spcPct
                }
                else
                {
                    _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", pr.SpaceAfter.ToString());
                    _writer.WriteEndElement(); //spcPct
                }
                _writer.WriteEndElement(); //spcAft
            }

            if (pr9 != null)
            if (pr9.BulletBlipReferencePresent)
            {
                foreach (ProgTags progtags in _ctx.Ppt.DocumentRecord.FirstChildWithType<List>().AllChildrenWithType<ProgTags>())
                {
                    foreach (ProgBinaryTag bintags in progtags.AllChildrenWithType<ProgBinaryTag>())
                    {
                        foreach (ProgBinaryTagDataBlob data in bintags.AllChildrenWithType<ProgBinaryTagDataBlob>())
                        {
                            foreach (BlipCollection9Container blips in data.AllChildrenWithType<BlipCollection9Container>())
                            {
                                if (blips.Children.Count > pr9.bulletblipref)
                                {
                                    BitmapBlip b = ((BlipEntityAtom)blips.Children[pr9.bulletblipref]).blip;
                                    ImagePart imgPart = null;
                                    imgPart = _parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                                    imgPart.TargetDirectory = "..\\media";
                                    System.IO.Stream outStream = imgPart.GetStream();
                                    outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);

                                    _writer.WriteStartElement("a", "buBlip", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
                                    _writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, imgPart.RelIdToString);
                                    _writer.WriteEndElement(); //blip
                                    _writer.WriteEndElement(); //buBlip
                                }
                            }
                        }
                    }
                }

                
            }

            new CharacterRunPropsMapping(_ctx, _writer).Apply(cr, "defRPr", _Master);                    

            _writer.WriteEndElement(); //lvlXpPr
        }
    }
}
