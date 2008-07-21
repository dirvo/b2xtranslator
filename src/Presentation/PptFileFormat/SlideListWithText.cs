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
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4080)]
    public class SlideListWithText : RegularContainer
    {
        public enum Instances
        {
            CollectionOfSlides = 0,
            CollectionOfMasterSlides = 1,
            CollectionOfNotesSlides = 2
        };

        /// <summary>
        /// List of all SlidePersistAtoms of this SlideListWithText.
        /// </summary>
        public List<SlidePersistAtom> SlidePersistAtoms = new List<SlidePersistAtom>();

        /// <summary>
        /// This dictionary manages a list of TextHeaderAtoms for each SlidePersistAtom.
        /// 
        /// Text of placeholders can appear in the SlideListWithText record.
        /// This dictionary is used for associating such text records with the slide they appear on.
        /// </summary>
        public Dictionary<SlidePersistAtom, List<TextHeaderAtom>> SlideToPlaceholderTextHeaders =
            new Dictionary<SlidePersistAtom,List<TextHeaderAtom>>();

        public SlideListWithText(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            SlidePersistAtom curSpAtom = null;
            TextHeaderAtom curThAtom = null;

            foreach (Record r in this.Children)
            {
                SlidePersistAtom spAtom = r as SlidePersistAtom;
                TextHeaderAtom thAtom = r as TextHeaderAtom;
                ITextDataRecord tdRecord = r as ITextDataRecord;

                if (spAtom != null)
                {
                    curSpAtom = spAtom;
                    this.SlidePersistAtoms.Add(spAtom);
                }

                else if (thAtom != null)
                {
                    curThAtom = thAtom;

                    if (!this.SlideToPlaceholderTextHeaders.ContainsKey(curSpAtom))
                        this.SlideToPlaceholderTextHeaders[curSpAtom] = new List<TextHeaderAtom>();

                    this.SlideToPlaceholderTextHeaders[curSpAtom].Add(thAtom);
                }

                else if (tdRecord != null)
                {
                    curThAtom.HandleTextDataRecord(tdRecord);
                }
            }
        }

        public TextHeaderAtom FindTextHeaderForOutlineTextRef(OutlineTextRefAtom otrAtom)
        {
            Slide slide = otrAtom.FirstAncestorWithType<Slide>();

            if (slide == null)
                throw new NotSupportedException("Can't find TextHeaderAtom for OutlineTextRefAtom which has no Slide ancestor");

            List<TextHeaderAtom> thAtoms = this.SlideToPlaceholderTextHeaders[slide.PersistAtom];
            return thAtoms[otrAtom.Index];
        }
    }

}
