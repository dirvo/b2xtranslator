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
    class AnimationMapping :
        AbstractOpenXmlMapping//,
        //IMapping<Dictionary<AnimationInfoContainer,int>>
    {
        protected ConversionContext _ctx;

        public AnimationMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public void Apply(Dictionary<AnimationInfoContainer, int> animations)
        {
            Dictionary<AnimationInfoAtom, int> animAtoms = new Dictionary<AnimationInfoAtom, int>();
            foreach (AnimationInfoContainer container in animations.Keys)
	        {
        		AnimationInfoAtom anim = container.FirstChildWithType<AnimationInfoAtom>();

                animAtoms.Add(anim, animations[container]);
                
            }
            writeTiming(animAtoms);
        }

        private int lastID = 0;
        private void writeTiming(Dictionary<AnimationInfoAtom, int> blindAtoms)
        {
            lastID = 0;

            _writer.WriteStartElement("p", "timing", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "tnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("dur", "indefinite");
            _writer.WriteAttributeString("restart", "never");
            _writer.WriteAttributeString("nodeType", "tmRoot");

            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "seq", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("concurrent", "1");
            _writer.WriteAttributeString("nextAc", "seek");

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("dur", "indefinite");
            _writer.WriteAttributeString("nodeType", "mainSeq");

            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            foreach (AnimationInfoAtom animinfo in blindAtoms.Keys)
            {
                writePar(animinfo, blindAtoms[animinfo].ToString());
            }
            //writePar(ShapeID);

            _writer.WriteEndElement(); //childTnLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteStartElement("p", "prevCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("evt", "onPrev");
            _writer.WriteAttributeString("delay", "0");

            _writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            _writer.WriteElementString("p", "sldTgt", OpenXmlNamespaces.PresentationML, "");

            _writer.WriteEndElement(); //tgtEl

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //prevCondLst

            _writer.WriteStartElement("p", "nextCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("evt", "onNext");
            _writer.WriteAttributeString("delay", "0");

            _writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            _writer.WriteElementString("p", "sldTgt", OpenXmlNamespaces.PresentationML, "");

            _writer.WriteEndElement(); //tgtEl

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //nextCondLst

            _writer.WriteEndElement(); //seq

            _writer.WriteEndElement(); //childTnLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteEndElement(); //par

            _writer.WriteEndElement(); //tnLst

            _writer.WriteStartElement("p", "bldLst", OpenXmlNamespaces.PresentationML);

            foreach (AnimationInfoAtom animinfo in blindAtoms.Keys)
            {
                _writer.WriteStartElement("p", "bldP", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("spid", blindAtoms[animinfo].ToString());
                _writer.WriteAttributeString("grpId", "0");
                _writer.WriteEndElement(); //bldP
            }            
           
            _writer.WriteEndElement(); //bldLst

            _writer.WriteEndElement(); //timing
        }


        private void writePar(AnimationInfoAtom animinfo, string ShapeID)
        {
            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("fill", "hold");

            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("delay", "0");

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //stCondLst

            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("fill", "hold");

            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("delay", "0");

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //stCondLst

            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("presetID", "3");
            _writer.WriteAttributeString("presetClass", "entr");
            _writer.WriteAttributeString("presetSubtype", "10");
            _writer.WriteAttributeString("fill", "hold");
            _writer.WriteAttributeString("grpId", "0");
            _writer.WriteAttributeString("nodeType", "clickEffect");

            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("delay", "0");

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //stCondLst

            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "set", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("dur", "1");
            _writer.WriteAttributeString("fill", "hold");

            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("delay", "0");

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //stCondLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("spid", ShapeID);

            _writer.WriteEndElement(); //spTgt

            _writer.WriteEndElement(); //tgtEl

            _writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, "style.visibility");

            _writer.WriteEndElement(); //attrNameLst

            _writer.WriteEndElement(); //cBhvr

            _writer.WriteStartElement("p", "to", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "strVal", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("val", "visible");

            _writer.WriteEndElement(); //str

            _writer.WriteEndElement(); //to

            _writer.WriteEndElement(); //set

            _writer.WriteStartElement("p", "animEffect", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("transition", "in");

            switch (animinfo.animEffect)
            {
                case 0x00: //Cut
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //not through black
                        case 0x02: //same as 0x00
                            _writer.WriteAttributeString("filter", "cut(false)");
                            break;
                        case 0x01: //through black
                            _writer.WriteAttributeString("filter", "cut(true)");
                            break;
                    }
                    break;
                case 0x01: //Random
                    _writer.WriteAttributeString("filter", "random");
                    break;
                case 0x02: //Blinds
                    if (animinfo.animEffectDirection == 0x01)
                    {
                        _writer.WriteAttributeString("filter", "blinds(horizontal)");
                    }
                    else
                    {
                        _writer.WriteAttributeString("filter", "blinds(vertical)");
                    }
                    break;
                case 0x03: //Checker
                    if (animinfo.animEffectDirection == 0x01)
                    {
                        _writer.WriteAttributeString("filter", "checker(horz)");
                    }
                    else
                    {
                        _writer.WriteAttributeString("filter", "checker(vert)");
                    }
                    break;                   
                case 0x04: //Cover
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //r->l
                            _writer.WriteAttributeString("filter", "cover(l)");
                            break;
                        case 0x01: //b->t
                            _writer.WriteAttributeString("filter", "cover(u)");
                            break;
                        case 0x02: //l->r
                            _writer.WriteAttributeString("filter", "cover(r)");
                            break;
                        case 0x03: //t->b
                            _writer.WriteAttributeString("filter", "cover(d)");
                            break;
                        case 0x04: //br->tl
                            _writer.WriteAttributeString("filter", "cover(lu)");
                            break;
                        case 0x05: //bl->tr
                            _writer.WriteAttributeString("filter", "cover(ru)");
                            break;
                        case 0x06: //tr->bl
                            _writer.WriteAttributeString("filter", "cover(ld)");
                            break;
                        case 0x07: //tl->br
                            _writer.WriteAttributeString("filter", "cover(rd)");
                            break;
                    }
                    break;
                case 0x05: //Dissolve
                    _writer.WriteAttributeString("filter", "dissolve");
                    break;
                case 0x06: //Fade
                    _writer.WriteAttributeString("filter", "fade");
                    break;
                case 0x07: //Pull
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //r->l
                            _writer.WriteAttributeString("filter", "pull(l)");
                            break;
                        case 0x01: //b->t
                            _writer.WriteAttributeString("filter", "pull(u)");
                            break;
                        case 0x02: //l->r
                            _writer.WriteAttributeString("filter", "pull(r)");
                            break;
                        case 0x03: //t->b
                            _writer.WriteAttributeString("filter", "pull(d)");
                            break;
                        case 0x04: //br->tl
                            _writer.WriteAttributeString("filter", "pull(lu)");
                            break;
                        case 0x05: //bl->tr
                            _writer.WriteAttributeString("filter", "pull(ru)");
                            break;
                        case 0x06: //tr->bl
                            _writer.WriteAttributeString("filter", "pull(ld)");
                            break;
                        case 0x07: //tl->br
                            _writer.WriteAttributeString("filter", "pull(rd)");
                            break;
                    }
                    break;
                case 0x08: //Random bar
                    if (animinfo.animEffectDirection == 0x01)
                    {
                        _writer.WriteAttributeString("filter", "randomBar(horz)");
                    }
                    else
                    {
                        _writer.WriteAttributeString("filter", "randomBar(vert)");
                    }
                    break;          
                case 0x09: //Strips
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x04: //br->tl
                            _writer.WriteAttributeString("filter", "strips(lu)");
                            break;
                        case 0x05: //bl->tr
                            _writer.WriteAttributeString("filter", "strips(ru)");
                            break;
                        case 0x06: //tr->bl
                            _writer.WriteAttributeString("filter", "strips(ld)");
                            break;
                        case 0x07: //tl->br
                            _writer.WriteAttributeString("filter", "strips(rd)");
                            break;
                    }
                    break;
                case 0x0a: //Wipe
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //r->l
                            _writer.WriteAttributeString("filter", "wipe(l)");
                            break;
                        case 0x01: //b->t
                            _writer.WriteAttributeString("filter", "wipe(u)");
                            break;
                        case 0x02: //l->r
                            _writer.WriteAttributeString("filter", "wipe(r)");
                            break;
                        case 0x03: //t->b
                            _writer.WriteAttributeString("filter", "wipe(d)");
                            break;
                    }
                    break;
                case 0x0b: //Zoom (box)
                    if (animinfo.animEffectDirection == 0x00)
                    {
                        _writer.WriteAttributeString("filter", "box(out)");
                    }
                    else
                    {
                        _writer.WriteAttributeString("filter", "box(in)");
                    }
                    break;
                case 0x0c: //Fly
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //from left
                            _writer.WriteAttributeString("filter", "fly(l)");
                            break;
                        case 0x01: //from top
                            _writer.WriteAttributeString("filter", "fly(u)");
                            break;
                        case 0x02: //from right
                            _writer.WriteAttributeString("filter", "fly(r)");
                            break;
                        case 0x03: //from bottom
                            _writer.WriteAttributeString("filter", "fly(d)");
                            break;
                        case 0x04: //from top left
                            _writer.WriteAttributeString("filter", "fly(lu)");
                            break;
                        case 0x05: //from top right
                            _writer.WriteAttributeString("filter", "fly(ru)");
                            break;
                        case 0x06: //from bottom left
                            _writer.WriteAttributeString("filter", "fly(ld)");
                            break;
                        case 0x07: //from bottom right
                            _writer.WriteAttributeString("filter", "fly(rd)");
                            break;
                        case 0x08: //from left edge of shape / text
                            _writer.WriteAttributeString("filter", "fly(l)");
                            break;
                        case 0x09: //from bottom edge of shape / text
                            _writer.WriteAttributeString("filter", "fly(u)");
                            break;
                        case 0x0a: //from right edge of shape / text
                            _writer.WriteAttributeString("filter", "fly(r)");
                            break;
                        case 0x0b: //from top edge of shape / text
                            _writer.WriteAttributeString("filter", "fly(d)");
                            break;
                        case 0x0c: //crawl from left
                            _writer.WriteAttributeString("filter", "fly(lu)");
                            break;
                        case 0x0d: //crawl from top 
                            _writer.WriteAttributeString("filter", "fly(ru)");
                            break;
                        case 0x0e: //crawl from right
                            _writer.WriteAttributeString("filter", "fly(ld)");
                            break;
                        case 0x0f: //crawl from bottom
                            _writer.WriteAttributeString("filter", "fly(rd)");
                            break;
                        case 0x10: //zoom 0 -> 1
                            _writer.WriteAttributeString("filter", "fly(rd)");
                            break;
                        //case 0x0f: //crawl from bottom
                        //    _writer.WriteAttributeString("filter", "fly(rd)");
                        //    break;
                        //case 0x0f: //crawl from bottom
                        //    _writer.WriteAttributeString("filter", "fly(rd)");
                        //    break;
                        //case 0x0f: //crawl from bottom
                        //    _writer.WriteAttributeString("filter", "fly(rd)");
                        //    break;
                    }
                    break;
                case 0x0d: //Split
                case 0x0e: //Flash
                case 0x0f:
                case 0x11: //Diamond
                case 0x12: //Plus
                case 0x13: //Wedge
                case 0x14:
                case 0x15:
                case 0x16:
                case 0x17:
                case 0x18:
                case 0x19:
                case 0x1a: //Wheel
                case 0x1b: //Circle
                default:
                    _writer.WriteAttributeString("filter", "blinds(horizontal)");
                    break;
            }

            _writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("dur", "500");
            _writer.WriteEndElement(); //cTn

            _writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("spid", ShapeID);

            _writer.WriteEndElement(); //spTgt

            _writer.WriteEndElement(); //tgtEl

            _writer.WriteEndElement(); //cBhvr

            _writer.WriteEndElement(); //animEffect

            _writer.WriteEndElement(); //childTnLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteEndElement(); //par

            _writer.WriteEndElement(); //childTnLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteEndElement(); //par

            _writer.WriteEndElement(); //childTnLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteEndElement(); //par
        }

    }
}
