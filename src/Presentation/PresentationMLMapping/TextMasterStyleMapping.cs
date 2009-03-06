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
        RegularContainer _Master;
        public PresentationMapping<RegularContainer> _parentSlideMapping = null;

        private int lastSpaceBefore = 0;
        private string lastColor = "";
        private string lastBulletFont = "";
        private string lastSize = "";
        public TextMasterStyleMapping(ConversionContext ctx, XmlWriter writer, PresentationMapping<RegularContainer> parentSlideMapping)
            : base(writer)
        {
            _ctx = ctx;
            _parentSlideMapping = parentSlideMapping;
        }

        public List<TextMasterStyleAtom> titleAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> bodyAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> CenterBodyAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> CenterTitleAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> noteAtoms = new List<TextMasterStyleAtom>();
        public void Apply(RegularContainer Master)
        {
            _Master = Master;

            List<TextMasterStyleAtom> atoms = Master.AllChildrenWithType<TextMasterStyleAtom>();            

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
            
            foreach (TextMasterStyleAtom atom in atoms)
            {
                if (atom.Instance == 0) titleAtoms.Add(atom);   
                if (atom.Instance == 1) bodyAtoms.Add(atom);
                if (atom.Instance == 5) CenterBodyAtoms.Add(atom);
                if (atom.Instance == 6) CenterTitleAtoms.Add(atom);
            }
            
            _writer.WriteStartElement("p", "txStyles", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "titleStyle", OpenXmlNamespaces.PresentationML);

            ParagraphRun9 pr9 = null;
            foreach (TextMasterStyleAtom atom in titleAtoms)
            {
                lastSpaceBefore = 0;
                lastBulletFont = "";
                lastColor = "";
                lastSize = "";
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    if (title9atoms.Count > 0 && title9atoms[0].pruns.Count > i) pr9 = title9atoms[0].pruns[i];
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i, true);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    if (title9atoms.Count > 0 && title9atoms[0].pruns.Count > i) pr9 = title9atoms[0].pruns[i];
                    writepPr(atom.CRuns[0], atom.PRuns[0], pr9, i, true);
                }
            }

            _writer.WriteEndElement(); //titleStyle

            _writer.WriteStartElement("p", "bodyStyle", OpenXmlNamespaces.PresentationML);

            foreach (TextMasterStyleAtom atom in bodyAtoms)
            {
                lastSpaceBefore = 0;
                lastColor = "";
                lastBulletFont = "";
                lastSize = "";
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    if (body9atoms.Count > 0 && body9atoms[0].pruns.Count > i) pr9 = body9atoms[0].pruns[i];
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i, false);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    if (body9atoms.Count > 0 && body9atoms[0].pruns.Count > i) pr9 = body9atoms[0].pruns[i];
                    writepPr(atom.CRuns[0], atom.PRuns[0],pr9, i, false);
                }
            }

            _writer.WriteEndElement(); //bodyStyle

            _writer.WriteEndElement(); //txStyles
        }

        public void ApplyNotesMaster(RegularContainer notesMaster)
        {
            _Master = notesMaster;
            MainMaster m = this._ctx.Ppt.MainMasterRecords[0];
            List<TextMasterStyleAtom> atoms = m.AllChildrenWithType<TextMasterStyleAtom>();
            foreach (TextMasterStyleAtom atom in atoms)
            {
                if (atom.Instance == 2) noteAtoms.Add(atom);
            }

            _writer.WriteStartElement("p", "notesStyle", OpenXmlNamespaces.PresentationML);

            ParagraphRun9 pr9 = null;
            foreach (TextMasterStyleAtom atom in noteAtoms)
            {
                lastSpaceBefore = 0;
                lastBulletFont = "";
                lastColor = "";
                lastSize = "";
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i, true);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    writepPr(atom.CRuns[0], atom.PRuns[0], pr9, i, true);
                }
            }

            _writer.WriteEndElement();
        }

        private void writepPr(CharacterRun cr, ParagraphRun pr, ParagraphRun9 pr9, int IndentLevel, bool isTitle)
        {
          
            //TextMasterStyleAtom defaultStyle = _ctx.Ppt.DocumentRecord.FirstChildWithType<DIaLOGIKa.b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<TextMasterStyleAtom>();
            
            _writer.WriteStartElement("a", "lvl" + (IndentLevel+1).ToString() + "pPr", OpenXmlNamespaces.DrawingML);

            //marL
            if (pr.LeftMarginPresent) _writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)pr.LeftMargin).ToString());
            //marR
            //lvl
            if (pr.IndentLevel > 0) _writer.WriteAttributeString("lvl", pr.IndentLevel.ToString());
            //indent
            if (pr.IndentPresent) _writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(pr.LeftMargin - pr.Indent)))).ToString());
            //algn
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
            //defTabSz
            //rtl
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
            //eaLnkBrk
            //fontAlgn
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
            //latinLnBrk
            //hangingPunct


            //lnSpc
            //spcBef
            if (pr.SpaceBeforePresent)
            {
                _writer.WriteStartElement("a", "spcBef", OpenXmlNamespaces.DrawingML);
                if (pr.SpaceBefore < 0)
                {
                    _writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", (-1 * 12 * pr.SpaceBefore).ToString()); //TODO: the 12 is wrong
                    _writer.WriteEndElement(); //spcPct
                }
                else
                {
                    _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    _writer.WriteAttributeString("val", (1000 * pr.SpaceBefore).ToString());
                    _writer.WriteEndElement(); //spcPct
                }
                _writer.WriteEndElement(); //spcBef
                lastSpaceBefore = (int)pr.SpaceBefore;
            }
            else
            {
                if (lastSpaceBefore != 0)
                {
                    _writer.WriteStartElement("a", "spcBef", OpenXmlNamespaces.DrawingML);
                    if (lastSpaceBefore < 0)
                    {
                        _writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (-1 * 12 * lastSpaceBefore).ToString()); //TODO: the 12 is wrong
                        _writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        _writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", (1000 * lastSpaceBefore).ToString());
                        _writer.WriteEndElement(); //spcPct
                    }
                    _writer.WriteEndElement(); //spcBef
                }
            }
            //spcAft
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
            //EG_TextBulletColor
            //EG_TextBulletSize
            //EG_TextBulletTypeFace
            //EG_TextBullet


            bool bulletwritten = false;
            if (pr9 != null)
            {
                if (pr9.BulletBlipReferencePresent)
                    foreach (ProgTags progtags in _ctx.Ppt.DocumentRecord.FirstChildWithType<List>().AllChildrenWithType<ProgTags>())
                    {
                        foreach (ProgBinaryTag bintags in progtags.AllChildrenWithType<ProgBinaryTag>())
                        {
                            foreach (ProgBinaryTagDataBlob data in bintags.AllChildrenWithType<ProgBinaryTagDataBlob>())
                            {
                                foreach (BlipCollection9Container blips in data.AllChildrenWithType<BlipCollection9Container>())
                                {
                                    if (blips.Children.Count > pr9.bulletblipref & pr9.bulletblipref > -1)
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
                                        bulletwritten = true;
                                    }
                                }
                            }
                        }
                    }
            }

            if (!bulletwritten & !isTitle)
            {
                if (pr.BulletFlagsFieldPresent & (pr.BulletFlags & (ushort)ParagraphMask.HasBullet) == 0)
                {
                    _writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");
                }
                else
                {
                    if (pr.BulletSizePresent)
                    {
                        if (pr.BulletSize > 0)
                        {
                            _writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", (pr.BulletSize * 1000).ToString());
                            _writer.WriteEndElement(); //buChar
                        }
                        else
                        {
                            //TODO
                        }
                     }
                     if (pr.BulletFontPresent)
                     {
                        _writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                        FontCollection fonts = _ctx.Ppt.DocumentRecord.FirstChildWithType<DIaLOGIKa.b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                        FontEntityAtom entity = fonts.entities[(int)pr.BulletTypefaceIdx];
                        if (entity.TypeFace.IndexOf('\0') > 0)
                        {
                            _writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                        }
                        else
                        {
                            _writer.WriteAttributeString("typeface", entity.TypeFace);
                        }
                        _writer.WriteEndElement(); //buChar
                        lastBulletFont = entity.TypeFace;
                     }
                     else if (lastBulletFont.Length > 0)
                     {
                         _writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                         if (lastBulletFont.IndexOf('\0') > 0)
                         {
                             _writer.WriteAttributeString("typeface", lastBulletFont.Substring(0, lastBulletFont.IndexOf('\0')));
                         }
                         else
                         {
                             _writer.WriteAttributeString("typeface", lastBulletFont);
                         }
                         _writer.WriteEndElement(); //buChar
                     }
                     if (pr.BulletCharPresent)
                     {
                        _writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("char", pr.BulletChar.ToString());
                        _writer.WriteEndElement(); //buChar
                     }
                 }
            }
            
            //tabLst
            //defRPr
            //extLst

            new CharacterRunPropsMapping(_ctx, _writer).Apply(cr, "defRPr", (RegularContainer)_Master, ref lastColor, ref lastSize);                    

            _writer.WriteEndElement(); //lvlXpPr
        }
    }
}
