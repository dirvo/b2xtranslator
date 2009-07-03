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
using System.Drawing;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class AnimationMapping :
        AbstractOpenXmlMapping//,
        //IMapping<Dictionary<AnimationInfoContainer,int>>
    {
        protected ConversionContext _ctx;
        private ShapeTreeMapping _stm;
        private List<Point> TextAreasForAnimation = new List<Point>();

        public AnimationMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public void Apply(SlideShowSlideInfoAtom slideshow)
        {
            if (slideshow.fAutoAdvance)
            {
                _writer.WriteStartElement("p", "transition", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("advTm", slideshow.slideTime.ToString());
                _writer.WriteEndElement();
            }
        }

        public void Apply(ProgBinaryTagDataBlob blob, PresentationMapping<RegularContainer> parentMapping, Dictionary<AnimationInfoContainer, int> animations, ShapeTreeMapping stm)
        {
            _parentMapping = parentMapping;
            _stm = stm;
            Dictionary<AnimationInfoAtom, int> animAtoms = new Dictionary<AnimationInfoAtom, int>();
            foreach (AnimationInfoContainer container in animations.Keys)
            {
                AnimationInfoAtom anim = container.FirstChildWithType<AnimationInfoAtom>();
                animAtoms.Add(anim, animations[container]);
            }

             ExtTimeNodeContainer c1 = blob.FirstChildWithType<ExtTimeNodeContainer>();
             //if (animAtoms.Count > 0)
             //{
             //    writeTiming(animAtoms, blob);
             //}
             //else
             //{
                 if (c1 != null)
                 {
                     ExtTimeNodeContainer c2 = c1.FirstChildWithType<ExtTimeNodeContainer>();
                     if (c2 != null)
                     {
                         ExtTimeNodeContainer c3 = c2.FirstChildWithType<ExtTimeNodeContainer>();
                         if (c3 != null)
                         {
                             writeTiming(animAtoms, blob);
                         }
                     }
                 }
             //}
        }

        private PresentationMapping<RegularContainer> _parentMapping;
        //public void Apply(ProgBinaryTagDataBlob blob, SlideMapping parentMapping)
        //{
        //    _parentMapping = parentMapping;
        //    Dictionary<AnimationInfoAtom, int> animAtoms = new Dictionary<AnimationInfoAtom, int>();
        //    writeTiming(animAtoms, blob);
        //}

        private VisualShapeAtom getShapeID(ExtTimeNodeContainer c)
        {
            List<VisualShapeAtom> lst = new List<VisualShapeAtom>();

            foreach (ExtTimeNodeContainer c8 in c.AllChildrenWithType<ExtTimeNodeContainer>())
                foreach (ExtTimeNodeContainer c9 in c8.AllChildrenWithType<ExtTimeNodeContainer>())
                {
                    foreach (TimeEffectBehaviorContainer c10a in c9.AllChildrenWithType<TimeEffectBehaviorContainer>())
                        foreach (TimeBehaviorContainer c10aa in c10a.AllChildrenWithType<TimeBehaviorContainer>())
                            foreach (ClientVisualElementContainer c10aaa in c10aa.AllChildrenWithType<ClientVisualElementContainer>())
                                foreach (VisualShapeAtom c10aaaa in c10aaa.AllChildrenWithType<VisualShapeAtom>())
                                    lst.Add(c10aaaa); //.shapeIdRef); //return c10aaaa.shapeIdRef;
                    foreach (TimeSetBehaviourContainer c10b in c9.AllChildrenWithType<TimeSetBehaviourContainer>())
                        foreach (TimeBehaviorContainer c10aa in c10b.AllChildrenWithType<TimeBehaviorContainer>())
                            foreach (ClientVisualElementContainer c10aaa in c10aa.AllChildrenWithType<ClientVisualElementContainer>())
                                foreach (VisualShapeAtom c10aaaa in c10aaa.AllChildrenWithType<VisualShapeAtom>())
                                    lst.Add(c10aaaa); //.shapeIdRef); //return c10aaaa.shapeIdRef;                     
                    foreach (TimeRotationBehaviorContainer r in c9.AllChildrenWithType<TimeRotationBehaviorContainer>())
                        foreach (TimeBehaviorContainer rr in r.AllChildrenWithType<TimeBehaviorContainer>())
                            foreach (ClientVisualElementContainer c10aaa in rr.AllChildrenWithType<ClientVisualElementContainer>())
                                foreach (VisualShapeAtom c10aaaa in c10aaa.AllChildrenWithType<VisualShapeAtom>())
                                    lst.Add(c10aaaa); //.shapeIdRef); //return c10aaaa.shapeIdRef;
                    foreach (TimeCommandBehaviorContainer r in c9.AllChildrenWithType<TimeCommandBehaviorContainer>())
                        foreach (TimeBehaviorContainer rr in r.AllChildrenWithType<TimeBehaviorContainer>())
                            foreach (ClientVisualElementContainer c10aaa in rr.AllChildrenWithType<ClientVisualElementContainer>())
                                foreach (VisualShapeAtom c10aaaa in c10aaa.AllChildrenWithType<VisualShapeAtom>())
                                    lst.Add(c10aaaa); //.shapeIdRef); //return c10aaaa.shapeIdRef;
                    foreach (TimeMotionBehaviorContainer r in c9.AllChildrenWithType<TimeMotionBehaviorContainer>())
                        foreach (TimeBehaviorContainer rr in r.AllChildrenWithType<TimeBehaviorContainer>())
                            foreach (ClientVisualElementContainer c10aaa in rr.AllChildrenWithType<ClientVisualElementContainer>())
                                foreach (VisualShapeAtom c10aaaa in c10aaa.AllChildrenWithType<VisualShapeAtom>())
                                    lst.Add(c10aaaa); //.shapeIdRef); //return c10aaaa.shapeIdRef;
                    foreach (TimeScaleBehaviorContainer c10a in c9.AllChildrenWithType<TimeScaleBehaviorContainer>())
                        foreach (TimeBehaviorContainer c10aa in c10a.AllChildrenWithType<TimeBehaviorContainer>())
                            foreach (ClientVisualElementContainer c10aaa in c10aa.AllChildrenWithType<ClientVisualElementContainer>())
                                foreach (VisualShapeAtom c10aaaa in c10aaa.AllChildrenWithType<VisualShapeAtom>())
                                    lst.Add(c10aaaa); //.shapeIdRef); //return c10aaaa.shapeIdRef;
                }

            return lst[0];
            //return 0;
        }

        //public void Apply(Dictionary<AnimationInfoContainer, int> animations)
        //{
        //    Dictionary<AnimationInfoAtom, int> animAtoms = new Dictionary<AnimationInfoAtom, int>();
        //    foreach (AnimationInfoContainer container in animations.Keys)
        //    {
        //        AnimationInfoAtom anim = container.FirstChildWithType<AnimationInfoAtom>();
        //        animAtoms.Add(anim, animations[container]);                
        //    }         
           
        //    writeTiming(animAtoms, null);
        //}

        private int lastID = 0;
        private void writeTiming(Dictionary<AnimationInfoAtom, int> blindAtoms, ProgBinaryTagDataBlob blob)
        {
            lastID = 0;

            _writer.WriteStartElement("p", "timing", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "tnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("dur", "indefinite");
            //_writer.WriteAttributeString("restart", "never");
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

            //foreach (AnimationInfoAtom animinfo in blindAtoms.Keys)
            //{
            //    writePar(animinfo, blindAtoms[animinfo].ToString());
            //}
            if (blob != null)
            {

                ExtTimeNodeContainer c1 = blob.FirstChildWithType<ExtTimeNodeContainer>();
                if (c1 != null)
                {
                    ExtTimeNodeContainer c2 = c1.FirstChildWithType<ExtTimeNodeContainer>();
                    if (c2 != null)
                    {

                        //ExtTimeNodeContainer c3 = c2.FirstChildWithType<ExtTimeNodeContainer>();
                        foreach (ExtTimeNodeContainer c3 in c2.AllChildrenWithType<ExtTimeNodeContainer>())
                        if (c3 != null)
                        {

                            int counter = 0;
                            AnimationInfoAtom a;
                            System.Collections.Generic.List<AnimationInfoAtom> atoms = new List<AnimationInfoAtom>();
                            foreach (AnimationInfoAtom key in blindAtoms.Keys) atoms.Add(key);

                            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

                            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
                            _writer.WriteAttributeString("id", (++lastID).ToString());
                            _writer.WriteAttributeString("fill", "hold");

                            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);


                            foreach (TimeConditionContainer c in c3.AllChildrenWithType<TimeConditionContainer>())
                            {
                                TimeConditionAtom t = c.FirstChildWithType<TimeConditionAtom>();

                                _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

                                switch (t.triggerEvent)
                                {
                                    case 0x0: //none
                                        break;
                                    case 0x1: //onBegin
                                        _writer.WriteAttributeString("evt", "onBegin");
                                        break;
                                    case 0x3: //Start
                                        _writer.WriteAttributeString("evt", "begin");
                                        break;
                                    case 0x4: //End
                                        _writer.WriteAttributeString("evt", "end");
                                        break;
                                    case 0x5: //Mouse click
                                        _writer.WriteAttributeString("evt", "onClick");
                                        break;
                                    case 0x7: //Mouse over
                                        _writer.WriteAttributeString("evt", "onMouseOver");
                                        break;
                                    case 0x9: //OnNext
                                        _writer.WriteAttributeString("evt", "onNext");
                                        break;
                                    case 0xa: //OnPrev
                                        _writer.WriteAttributeString("evt", "onPrev");
                                        break;
                                    case 0xb: //Stop audio
                                        _writer.WriteAttributeString("evt", "onStopAudio");
                                        break;
                                    default:
                                        break;
                                }

                                if (t.delay == -1)
                                {
                                    _writer.WriteAttributeString("delay", "indefinite");
                                }
                                else
                                {
                                    _writer.WriteAttributeString("delay", t.delay.ToString());
                                }

                                if (t.triggerObject == TimeConditionAtom.TriggerObjectEnum.TimeNode)
                                {
                                    _writer.WriteStartElement("p", "tn", OpenXmlNamespaces.PresentationML);
                                    _writer.WriteAttributeString("val", t.id.ToString());
                                    _writer.WriteEndElement();
                                }

                                _writer.WriteEndElement(); //cond

                            }

                            _writer.WriteEndElement(); //stCondLst

                            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);


                            foreach (ExtTimeNodeContainer c4 in c3.AllChildrenWithType<ExtTimeNodeContainer>())
                            {
                                a = null;
                                if (atoms.Count > counter) a = atoms[counter];
                                writePar2(c4, a);
                                counter++;
                            }

                            _writer.WriteEndElement(); //childTnLst

                            _writer.WriteEndElement(); //cTn

                            _writer.WriteEndElement(); //par
                        }
                    }
                }
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

            if (blindAtoms.Count > 0)
            {

                _writer.WriteStartElement("p", "bldLst", OpenXmlNamespaces.PresentationML);

                foreach (AnimationInfoAtom animinfo in blindAtoms.Keys)
                {
                    _writer.WriteStartElement("p", "bldP", OpenXmlNamespaces.PresentationML);
                    _writer.WriteAttributeString("spid", blindAtoms[animinfo].ToString());
                    _writer.WriteAttributeString("grpId", "0");

                    if (animinfo.animBuildType == AnimationInfoAtom.AnimBuildTypeEnum.Level1Build) _writer.WriteAttributeString("build", "p");
                    if (animinfo.fAnimateBg) _writer.WriteAttributeString("animBg", "1");

                    _writer.WriteEndElement(); //bldP
                }

                _writer.WriteEndElement(); //bldLst

            }

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

            if (animinfo.fAutomatic)
            {
                _writer.WriteAttributeString("delay", "0");
            }
            else
            {
                _writer.WriteAttributeString("delay", "indefinite");
            }

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
            _writer.WriteAttributeString("presetID", (animinfo.animEffect +1).ToString()); //3
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

            if (true) //TODO: when?
            {
                writeAnimEffect(animinfo, ShapeID,-1);
            }
            else
            {
                writeFlyAnim(animinfo, ShapeID,-1);
            }

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

        private void writePar2(ExtTimeNodeContainer container, AnimationInfoAtom animinfo)
        {
          

            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("fill", "hold");

            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            if (container.FirstChildWithType<TimeConditionContainer>() != null)
            {
                _writer.WriteAttributeString("delay", container.FirstChildWithType<TimeConditionContainer>().FirstChildWithType<TimeConditionAtom>().delay.ToString());
            }
            else
            {
                _writer.WriteAttributeString("delay", "0");
            }

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //stCondLst

            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());

            string filter = "";
            foreach (ExtTimeNodeContainer c2 in container.AllChildrenWithType<ExtTimeNodeContainer>())
            {
                foreach (ExtTimeNodeContainer c3 in c2.AllChildrenWithType<ExtTimeNodeContainer>())
                {
                    foreach (TimeEffectBehaviorContainer c4 in c3.AllChildrenWithType<TimeEffectBehaviorContainer>())
                    {
                        foreach (TimeVariantValue v in c4.AllChildrenWithType<TimeVariantValue>())
                        {
                            if (v.type == TimeVariantTypeEnum.String)
                            {
                                filter = v.stringValue;
                            }
                        }
                    }
                }
            }

            if (animinfo != null)
            {
                _writer.WriteAttributeString("presetID", (animinfo.animEffect + 1).ToString()); //3
            }
            else
            {
                _writer.WriteAttributeString("presetID", "12"); //3
            }
            _writer.WriteAttributeString("presetClass", "entr");
            _writer.WriteAttributeString("presetSubtype", "4");
            _writer.WriteAttributeString("fill", "hold");
            //_writer.WriteAttributeString("grpId", "0");

            bool nodeTypeWritten = false;
            if (container.FirstChildWithType<ExtTimeNodeContainer>() != null)
            {
                ExtTimeNodeContainer c2 = container.FirstChildWithType<ExtTimeNodeContainer>();
                if (c2.FirstChildWithType<TimePropertyList4TimeNodeContainer>() != null)
                {
                    TimePropertyList4TimeNodeContainer c3 = c2.FirstChildWithType<TimePropertyList4TimeNodeContainer>();
                    TimeVariantValue v = c3.FirstChildWithType<TimeVariantValue>();

                    switch (v.intValue)
                    {
                        case 1:
                            _writer.WriteAttributeString("nodeType", "clickEffect");
                            nodeTypeWritten = true;
                            break;
                        case 2:
                            nodeTypeWritten = true;
                            _writer.WriteAttributeString("nodeType", "withEffect");
                            break;
                        case 3:
                            _writer.WriteAttributeString("nodeType", "afterEffect");
                            nodeTypeWritten = true;
                            break;
                        case 4:
                        case 5:
                        case 6:
                        case 7: 
                        case 8:
                        case 9:
                        default:
                            break;
                    }

                }

            }
            if (!nodeTypeWritten) _writer.WriteAttributeString("nodeType", "clickEffect");

            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("delay", "0");

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //stCondLst

            _writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "set", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("additive", "repl");

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            _writer.WriteAttributeString("dur", "1000");
            _writer.WriteAttributeString("fill", "hold");

            _writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            _writer.WriteAttributeString("delay", "0");

            _writer.WriteEndElement(); //cond

            _writer.WriteEndElement(); //stCondLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            VisualShapeAtom vsa = getShapeID(container);
            uint c4Id = vsa.shapeIdRef;
            string ShapeID = "";

            if (_stm.spidToId.ContainsKey((int)c4Id))
            {
                ShapeID = _stm.spidToId[(int)c4Id].ToString();
            }
            else
            {
                foreach (int sId in _stm.spidToId.Keys)
                {
                    if (sId > 0)
                    {
                        ShapeID = _stm.spidToId[sId].ToString();
                        break;
                    }
                }
            }
            _writer.WriteAttributeString("spid", ShapeID);

            int targetRun = -1;
            if (vsa.type == TimeVisualElementEnum.TextRange)
            {
                int i = 0;
                foreach (Point p in TextAreasForAnimation)
                {
                    if (p.X <= vsa.data1 && p.Y >= vsa.data2)
                    {
                        targetRun = i;
                        break;
                    }
                    i++;
                }
                if (targetRun == -1)
                {
                    TextAreasForAnimation.Add(new Point(vsa.data1, vsa.data2));
                    targetRun = TextAreasForAnimation.Count-1;
                }

                _writer.WriteStartElement("p", "txEl", OpenXmlNamespaces.PresentationML);
                _writer.WriteStartElement("p", "pRg", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("st", targetRun.ToString());
                _writer.WriteAttributeString("end", targetRun.ToString());
                _writer.WriteEndElement(); //pRg
                _writer.WriteEndElement(); //txEl
            }


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

            if (filter.Length > 0)
            {
                _writer.WriteStartElement("p", "animEffect", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("transition", "in");
                _writer.WriteAttributeString("filter", filter);
                _writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("additive", "repl");
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
            }
            else
            {
                if (animinfo != null)
                    if (animinfo.animEffect == 0x0c && animinfo.animEffectDirection > 0x3)
                    {
                        writeFlyAnim(animinfo, ShapeID, targetRun);
                    }
                    else
                    {
                        writeAnimEffect(animinfo, ShapeID, targetRun);
                    }
                
            }

           

            _writer.WriteEndElement(); //childTnLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteEndElement(); //par

            _writer.WriteEndElement(); //childTnLst

            _writer.WriteEndElement(); //cTn

            _writer.WriteEndElement(); //par

       }

        

        public void writeFlyAnim(AnimationInfoAtom animinfo, string ShapeID, int targetRun)
        {
            //X
            _writer.WriteStartElement("p", "anim", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("calcmode", "lin");
            _writer.WriteAttributeString("valueType", "num");

            _writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("additive", "base");

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            switch(animinfo.animEffectDirection)
            {
                case 0xC:
                case 0xD:
                case 0xE:
                case 0xF:
                    _writer.WriteAttributeString("dur", "5000");
                    break;
                default:
                    _writer.WriteAttributeString("dur", "500");
                    break;
            }
            _writer.WriteAttributeString("fill", "hold");
            _writer.WriteEndElement(); //cTn

            _writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("spid", ShapeID);


            if (targetRun != -1)
            {
                _writer.WriteStartElement("p", "txEl", OpenXmlNamespaces.PresentationML);
                _writer.WriteStartElement("p", "pRg", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("st", targetRun.ToString());
                _writer.WriteAttributeString("end", targetRun.ToString());
                _writer.WriteEndElement(); //pRg
                _writer.WriteEndElement(); //txEl
            }

            _writer.WriteEndElement(); //spTgt
            _writer.WriteEndElement(); //tgtEl

            _writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);
            switch (animinfo.animEffectDirection)
            {
                case 0x10: case 0x11: case 0x12: case 0x13: case 0x14: case 0x15:
                    _writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, "ppt_w");
                    break;
                default:
                    _writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, "ppt_x");
                    break;
            }
            
            _writer.WriteEndElement(); //attrNameLst

            _writer.WriteEndElement(); //cBhvr

            _writer.WriteStartElement("p", "tavLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "tav", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("tm", "0");
            _writer.WriteStartElement("p", "val", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "strVal", OpenXmlNamespaces.PresentationML);

            switch (animinfo.animEffectDirection)
            {
                case 0x0: case 0x4: case 0x6:
                    _writer.WriteAttributeString("val", "0-#ppt_w/2");
                    break;
                case 0x2: case 0x5: case 0x7:
                    _writer.WriteAttributeString("val", "1+#ppt_w/2");
                    break;
                case 0x10: case 0x11: case 0x12: case 0x13: case 0x14: case 0x15: //zoom
                    _writer.WriteAttributeString("val", "0");
                    break;
                default:
                    _writer.WriteAttributeString("val", "#ppt_x");
                    break;
            }

            _writer.WriteEndElement(); //strVal
            _writer.WriteEndElement(); //val
            _writer.WriteEndElement(); //tav

            _writer.WriteStartElement("p", "tav", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("tm", "100000");
            _writer.WriteStartElement("p", "val", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "strVal", OpenXmlNamespaces.PresentationML);
            switch (animinfo.animEffectDirection)
            {
                case 0x10: case 0x11: case 0x12: case 0x13: case 0x14: case 0x15: //zoom
                    _writer.WriteAttributeString("val", "#ppt_w");
                    break;
                default:
                    _writer.WriteAttributeString("val", "#ppt_x");
                    break;
            }
            _writer.WriteEndElement(); //strVal
            _writer.WriteEndElement(); //val
            _writer.WriteEndElement(); //tav

            _writer.WriteEndElement(); //tavLst

            _writer.WriteEndElement(); //anim

            //Y
            _writer.WriteStartElement("p", "anim", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("calcmode", "lin");
            _writer.WriteAttributeString("valueType", "num");

            _writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("additive", "base");

            _writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++lastID).ToString());
            switch (animinfo.animEffectDirection)
            {
                case 0xC: case 0xD: case 0xE: case 0xF:
                    _writer.WriteAttributeString("dur", "5000");
                    break;
                default:
                    _writer.WriteAttributeString("dur", "500");
                    break;
            }
            _writer.WriteAttributeString("fill", "hold");
            _writer.WriteEndElement(); //cTn

            _writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("spid", ShapeID);

            if (targetRun != -1)
            {
                _writer.WriteStartElement("p", "txEl", OpenXmlNamespaces.PresentationML);
                _writer.WriteStartElement("p", "pRg", OpenXmlNamespaces.PresentationML);
                _writer.WriteAttributeString("st", targetRun.ToString());
                _writer.WriteAttributeString("end", targetRun.ToString());
                _writer.WriteEndElement(); //pRg
                _writer.WriteEndElement(); //txEl
            }

            _writer.WriteEndElement(); //spTgt
            _writer.WriteEndElement(); //tgtEl

            _writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);
            switch (animinfo.animEffectDirection)
            {
                case 0x10:
                case 0x11:
                case 0x12:
                case 0x13:
                case 0x14:
                case 0x15:
                    _writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, "ppt_h");
                    break;
                default:
                    _writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, "ppt_y");
                    break;
            }
            _writer.WriteEndElement(); //attrNameLst

            _writer.WriteEndElement(); //cBhvr

            _writer.WriteStartElement("p", "tavLst", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "tav", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("tm", "0");
            _writer.WriteStartElement("p", "val", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "strVal", OpenXmlNamespaces.PresentationML);

            switch (animinfo.animEffectDirection)
            {
                case 0x1: case 0x6: case 0x5: case 0xd: //top
                    _writer.WriteAttributeString("val", "0-#ppt_h/2");
                    break;
                case 0x3: case 0x4: case 0x7: case 0xf: //bottom
                    _writer.WriteAttributeString("val", "1+#ppt_h/2");
                    break;
                case 0x10: case 0x11: case 0x12: case 0x13: case 0x14: case 0x15: //zoom
                    _writer.WriteAttributeString("val", "0");
                    break;
                default:
                    _writer.WriteAttributeString("val", "#ppt_y");
                    break;
            }


            _writer.WriteEndElement(); //strVal
            _writer.WriteEndElement(); //val
            _writer.WriteEndElement(); //tav

            _writer.WriteStartElement("p", "tav", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("tm", "100000");
            _writer.WriteStartElement("p", "val", OpenXmlNamespaces.PresentationML);
            _writer.WriteStartElement("p", "strVal", OpenXmlNamespaces.PresentationML);

            switch (animinfo.animEffectDirection)
            {
                case 0x1:
                case 0x6:
                case 0x5:
                case 0xd: //top
                    _writer.WriteAttributeString("val", "0-#ppt_h/2");
                    break;
                case 0x3:
                case 0x4:
                case 0x7:
                case 0xf: //bottom
                    _writer.WriteAttributeString("val", "1+#ppt_h/2");
                    break;
                case 0x10:
                case 0x11:
                case 0x12:
                case 0x13:
                case 0x14:
                case 0x15: //zoom
                    _writer.WriteAttributeString("val", "#ppt_h");
                    break;
                default:
                    _writer.WriteAttributeString("val", "#ppt_y");
                    break;
            }

            _writer.WriteEndElement(); //strVal
            _writer.WriteEndElement(); //val
            _writer.WriteEndElement(); //tav

            _writer.WriteEndElement(); //tavLst

            _writer.WriteEndElement(); //anim
        }

        public void writeAnimEffect(AnimationInfoAtom animinfo, string ShapeID, int targetRun)
        {
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
                    if (animinfo.animEffectDirection == 0x00)
                    {
                        _writer.WriteAttributeString("filter", "checkerboard(across)");
                    }
                    else
                    {
                        _writer.WriteAttributeString("filter", "checkerboard(down)");
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
                            _writer.WriteAttributeString("filter", "slide(fromLeft)");
                            break;
                        case 0x01: //from top
                            _writer.WriteAttributeString("filter", "slide(fromTop)");
                            break;
                        case 0x02: //from right
                            _writer.WriteAttributeString("filter", "slide(fromRight)");
                            break;
                        case 0x03: //from bottom  
                            _writer.WriteAttributeString("filter", "slide(fromBottom)");
                            break;
                        case 0x04: //from top left
                        case 0x05: //from top right
                        case 0x06: //from bottom left
                        case 0x07: //from bottom right
                        case 0x08: //from left edge of shape / text
                        case 0x09: //from bottom edge of shape / text
                        case 0x0a: //from right edge of shape / text
                        case 0x0b: //from top edge of shape / text
                        case 0x0c: //crawl from left
                        case 0x0d: //crawl from top 
                        case 0x0e: //crawl from right
                        case 0x0f: //crawl from bottom
                        case 0x10: //zoom 0 -> 1
                        case 0x11: //zoom 0.5 -> 1
                        case 0x12: //zoom 4 -> 1
                        case 0x13: //zoom 1.5 -> 1
                        case 0x14: //zoom 0 -> 1; screen center -> actual center
                        case 0x15: //zoom 4 -> 1; bottom -> actual position
                        case 0x16: //stretch center -> l & r
                        case 0x17: //stretch l -> r
                        case 0x18: //stretch t -> b
                        case 0x19: //stretch r -> l
                        case 0x1a: //stretch b -> t
                        case 0x1b: //rotate around vertical axis that passes through its center
                        case 0x1c: //flies in a spiral
                            _writer.WriteAttributeString("filter", "slide(fromBottom)");
                            break;
                    }
                    break;
                case 0x0d: //Split
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //horz m -> tb
                            _writer.WriteAttributeString("filter", "split(outHorizontal)");
                            break;
                        case 0x01: //horz tb -> m
                            _writer.WriteAttributeString("filter", "split(inHorizontal)");
                            break;
                        case 0x02: //vert m -> lr
                            _writer.WriteAttributeString("filter", "split(outVertical)");
                            break;
                        case 0x03: //vert
                            _writer.WriteAttributeString("filter", "split(inVertical)");
                            break;
                    }
                    break;
                case 0x0e: //Flash
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //after short time
                        case 0x01: //after medium time
                        case 0x02: //after long time
                            break;
                    }
                    break;
                case 0x0f:
                case 0x11: //Diamond
                    _writer.WriteAttributeString("filter", "diamond(out)");
                    break;
                case 0x12: //Plus
                    _writer.WriteAttributeString("filter", "plus");
                    break;
                case 0x13: //Wedge
                    _writer.WriteAttributeString("filter", "wedge");
                    break;
                case 0x14:
                case 0x15:
                case 0x16:
                case 0x17:
                case 0x18:
                case 0x19:
                case 0x1a: //Wheel
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x01: //1 spoke
                            _writer.WriteAttributeString("filter", "wheel(1)");
                            break;
                        case 0x02: //2 spokes
                            _writer.WriteAttributeString("filter", "wheel(2)");
                            break;
                        case 0x03: //3 spokes
                            _writer.WriteAttributeString("filter", "wheel(3)");
                            break;
                        case 0x04: //4 spokes
                            _writer.WriteAttributeString("filter", "wheel(4)");
                            break;
                        case 0x08: //8 spokes
                            _writer.WriteAttributeString("filter", "wheel(8)");
                            break;
                    }
                    break;
                case 0x1b: //Circle
                    _writer.WriteAttributeString("filter", "circle");
                    break;
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
        }
    }
}
