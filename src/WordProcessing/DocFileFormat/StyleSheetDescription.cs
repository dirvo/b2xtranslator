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
using System.Collections;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class StyleSheetDescription : IVisitable
    {
        public enum StyleKind
        {
            paragraph = 1,
            character,
            table,
            list
        }

        public enum StyleIdentifier
        {
            Normal = 0,
            Heading1,
            Heading2,
            Heading3,
            Heading4,
            Heading5,
            Heading6,
            Heading7,
            Heading8,
            Heading9,
            Index1,
            Index2,
            Index3,
            Index4,
            Index5,
            Index6,
            Index7,
            Index8,
            Index9,
            TOC1,
            TOC2,
            TOC3,
            TOC4,
            TOC5,
            TOC6,
            TOC7,
            TOC8,
            TOC9,
            NormalIndent, 
            FootnoteText,
            AnnotationText,
            Header,
            Footer,
            IndexHeading,
            Caption,
            ToCaption,
            EnvelopeAddress,
            EnvelopeReturn,
            FootnoteReference,
            AnnotationReference,
            LineNumber,
            PageNumber,
            EndnoteReference,
            EndnoteText,
            TableOfAuthoring,
            Macro,
            TOAHeading,
            List,
            ListBullet,
            ListNumber,
            List2,
            List3,
            List4,
            List5,
            ListBullet2,
            ListBullet3,
            ListBullet4,
            ListBullet5,
            ListNumber2,
            ListNumber3,
            ListNumber4,
            ListNumber5,
            Title,
            Closing,
            Signature,
            NormalCharacter,
            BodyText,
            BodyTextIndent,
            ListContinue,
            ListContinue2,
            ListContinue3,
            ListContinue4,
            ListContinue5,
            MessageHeader,
            Subtitle,
            Salutation,
            Date,
            BodyText1I,
            BodyText1I2,
            NoteHeading,
            BodyText2,
            BodyText3,
            BodyTextIndent2,
            BodyTextIndent3,
            BlockQuote,
            Hyperlink,
            FollowedHyperlink,
            Strong,
            Emphasis,
            NavPane,
            PlainText,
            AutoSignature,
            FormTop,
            FormBottom,
            HtmlNormal,
            HtmlAcronym,
            HtmlAddress,
            HtmlCite,
            HtmlCode,
            HtmlDfn,
            HtmlKbd,
            HtmlPre,
            htmlSamp,
            HtmlTt,
            HtmlVar,
            TableNormal,
            AnnotationSubject,
            NormalList,
            OutlineList1,
            OutlineList2,
            OutlineList3,
            TableSimple,
            TableSimple2,
            TableSimple3,
            TableClassic1,
            TableClassic2,
            TableClassic3,
            TableClassic4,
            TableColorful1,
            TableColorful2,
            TableColorful3,
            TableColumns1,
            TableColumns2,
            TableColumns3,
            TableColumns4,
            TableColumns5,
            TableGrid1,
            TableGrid2,
            TableGrid3,
            TableGrid4,
            TableGrid5,
            TableGrid6,
            TableGrid7,
            TableGrid8,
            TableList1,
            TableList2,
            TableList3,
            TableList4,
            TableList5,
            TableList6,
            TableList7,
            TableList8,
            Table3DFx1,
            Table3DFx2,
            Table3DFx3,
            TableContemporary,
            TableElegant,
            TableProfessional,
            TableSubtle1,
            tableSubtle2,
            TableWeb1,
            TableWeb2,
            TableWeb3,
            Acetate,
            TableGrid,
            TableTheme,
            Max,
            User = 4094,
            Null = 4095
        }

        /// <summary>
        /// The name of the style
        /// </summary>
        public string xstzName;

        /// <summary>
        /// Invariant style identifier 
        /// </summary>
        public StyleIdentifier sti;

        /// <summary>
        /// spare field for any temporary use, always reset back to zero! 
        /// </summary>
        public bool fScratch;

        /// <summary>
        /// PHEs of all text with this style are wrong
        /// </summary>
        public bool fInvalHeight;

        /// <summary>
        /// UPEs have been generated 
        /// </summary>
        public bool fHasUpe;

        /// <summary>
        /// std has been mass-copied; if unused at save time, 
        /// style should be deleted 
        /// </summary>
        public bool fMassCopy;

        /// <summary>
        /// style kind 
        /// </summary>
        public StyleKind stk;

        /// <summary>
        /// base style 
        /// </summary>
        public StyleIdentifier istdBase;

        /// <summary>
        /// number of UPXs (and UPEs) 
        /// </summary>
        public UInt16 cupx;

        /// <summary>
        /// next style
        /// </summary>
        public StyleIdentifier istdNext;

        /// <summary>
        /// offset to end of upx's, start of upe's 
        /// </summary>
        public UInt16 bchUpe;

        /// <summary>
        /// auto redefine style when appropriate 
        /// </summary>
        public bool fAutoRedef;

        /// <summary>
        /// hidden from UI? 
        /// </summary>
        public bool fHidden;

        /// <summary>
        /// style already has valid sprmCRgLidX_80 in it 
        /// </summary>
        public bool f97LidsSet;

        /// <summary>
        /// if f97LidsSet, says whether we copied the lid from sprmCRgLidX 
        /// into sprmCRgLidX_80 or whether we gotrid of sprmCRgLidX_80
        /// </summary>
        public bool fCopyLang;

        /// <summary>
        /// HTML Threading compose style 
        /// </summary>
        public bool fPersonalCompose;

        /// <summary>
        /// HTML Threading reply style 
        /// </summary>
        public bool fPersonalReply;

        /// <summary>
        /// HTML Threading - another user's personal style 
        /// </summary>
        public bool fPersonal;

        /// <summary>
        /// Do not export this style to HTML/CSS 
        /// </summary>
        public bool fNoHtmlExport;

        /// <summary>
        /// Do not show this style in long style lists 
        /// </summary>
        public bool fSemiHidden;

        /// <summary>
        /// Locked style? 
        /// </summary>
        public bool fLocked;

        /// <summary>
        /// Style is used by a word feature, e.g. footnote
        /// </summary>
        public bool fInternalUse;

        /// <summary>
        /// Is this style linked to another?
        /// </summary>
        public StyleIdentifier istdLink;

        /// <summary>
        /// Style has RevMarking history 
        /// </summary>
        public bool fHasOriginalStyle;

        /// <summary>
        /// marks during merge which doc's style changed 
        /// </summary>
        public UInt32 rsid;

        /// <summary>
        /// used temporarily during html export 
        /// </summary>
        public UInt32 iftcHtml;

        /// <summary>
        /// A StyleSheetDescription can have a PAPX. <br/>
        /// If the style doesn't modify paragraph properties, papx is null.
        /// </summary>
        public ParagraphPropertyExceptions papx;

        /// <summary>
        /// A StyleSheetDescription can have a CHPX. <br/>
        /// If the style doesn't modify character properties, chpx is null.
        /// </summary>
        public CharacterPropertyExceptions chpx;

        /// <summary>
        /// A StyleSheetDescription can have a TAPX. <br/>
        /// If the style doesn't modify table properties, tapx is null.
        /// </summary>
        public TablePropertyExceptions tapx;

        public StyleSheetDescription(byte[] bytes)
        {
            BitArray bits = new BitArray(bytes);

            //parsing the fix part

            //sti
            BitArray stiBits = Utils.BitArrayCopy(bits, 0, 12);
            this.sti = (StyleIdentifier)Utils.BitArrayToUInt32(stiBits);
            //flags
            this.fScratch = bits[12];
            this.fInvalHeight = bits[13];
            this.fHasUpe = bits[14];
            this.fMassCopy = bits[15];
            //stk
            BitArray stkBits = Utils.BitArrayCopy(bits, 16, 4);
            this.stk = (StyleKind)Utils.BitArrayToUInt32(stkBits);
            //istdBase
            BitArray istdBits = Utils.BitArrayCopy(bits, 20, 12);
            this.istdBase = (StyleIdentifier)Utils.BitArrayToUInt32(istdBits);
            //cupx
            BitArray cupxBits = Utils.BitArrayCopy(bits, 32, 4);
            this.cupx = (UInt16)Utils.BitArrayToUInt32(cupxBits);
            //istdNext
            BitArray istdNextBits = Utils.BitArrayCopy(bits, 36, 12);
            this.istdNext = (StyleIdentifier)Utils.BitArrayToUInt32(istdNextBits);
            //bchUpe
            BitArray bchBits = Utils.BitArrayCopy(bits, 48, 16);
            this.bchUpe = (UInt16)Utils.BitArrayToUInt32(bchBits);
            //flags
            this.fAutoRedef = bits[64];
            this.fHidden = bits[65];
            this.f97LidsSet = bits[66];
            this.fCopyLang = bits[67];
            this.fPersonalCompose = bits[68];
            this.fPersonalReply = bits[69];
            this.fPersonal = bits[70];
            this.fNoHtmlExport = bits[71];
            this.fSemiHidden = bits[72];
            this.fLocked = bits[73];
            this.fInternalUse = bits[74];
            //istdLink
            BitArray istdLinkBits = Utils.BitArrayCopy(bits, 80, 12);
            this.istdLink = (StyleIdentifier)Utils.BitArrayToUInt32(istdLinkBits);
            //fHasOriginalStyle
            this.fHasOriginalStyle = bits[92];
            //rsid
            BitArray rsidBits = Utils.BitArrayCopy(bits, 96, 32);
            this.rsid = Utils.BitArrayToUInt32(rsidBits);

            //parsing the variable part

            //xstz
            byte characterCount = bytes[18];
            //characters are zero-terminated, so 1 char has 2 bytes:
            byte[] name = new byte[characterCount * 2];
            Array.Copy(bytes, 20, name, 0, name.Length);
            //remove zero-termination
            this.xstzName = Encoding.Unicode.GetString(name);

            //parse the UPX structs
            int upxOffset = 18 + 1+ (characterCount * 2) + 2;
            for (int i = 0; i < this.cupx; i++)
            {
                //find the next even byte border
                if (upxOffset % 2 != 0)
                {
                    upxOffset++;
                }

                //get the count of bytes for UPX
                UInt16 cbUPX = System.BitConverter.ToUInt16(bytes, upxOffset);

                if (cbUPX > 0)
                {
                    //copy the bytes
                    byte[] upxBytes = new byte[cbUPX];
                    Array.Copy(bytes, upxOffset + 2, upxBytes, 0, upxBytes.Length);

                    if (this.stk == StyleKind.table)
                    {
                        //first upx is TAPX; second PAPX, third CHPX
                        switch (i)
                        {
                            case 0: this.tapx = new TablePropertyExceptions(upxBytes); break;
                            case 1: this.papx = new ParagraphPropertyExceptions(upxBytes); break;
                            case 2: this.chpx = new CharacterPropertyExceptions(upxBytes); break;
                        }
                    }
                    else if (this.stk == StyleKind.paragraph)
                    {
                        //first upx is PAPX, second CHPX
                        switch (i)
                        {
                            case 0: 
                                this.papx = new ParagraphPropertyExceptions(upxBytes); 
                                break;
                            case 1: 
                                this.chpx = new CharacterPropertyExceptions(upxBytes); 
                                break;
                        }
                    }
                    else if (this.stk == StyleKind.list)
                    {
                        //list styles have only one PAPX
                        switch (i)
                        {
                            case 0: this.papx = new ParagraphPropertyExceptions(upxBytes); break;
                        }
                    }
                    else if (this.stk == StyleKind.character)
                    {
                        //character styles have only one CHPX
                        switch (i)
                        {
                            case 0: this.chpx = new CharacterPropertyExceptions(upxBytes); break;
                        }
                    }
                }

                //increase the offset for the next run
                upxOffset += (2 + cbUPX );
            }

        }

        /// <summary>
        /// Returns the style id.<br/>
        /// This will be the value of the enum, if it is a default style.
        /// </summary>
        /// <returns>The id</returns>
        public string getStyleIdentifier()
        {
            if (this.sti != StyleIdentifier.User)
            {
                return this.sti.ToString();
            }
            else
            {
                return this.xstzName.Replace(" ", "");
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<StyleSheetDescription>)mapping).Apply(this);
        }

        #endregion
    }
}
