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

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class SectionProperties
    {
        /// <summary>
        /// Break code:<br/>
        /// 0 No break<br/>
        /// 1 New column<br/>
        /// 2 New page<br/>
        /// 3 Even page<br/>
        /// 4 Odd page<br/>
        /// </summary>
        public byte bkc;

        /// <summary>
        /// Set to true whene a title page is to be displayed
        /// </summary>
        public bool fTitlePage;

        /// <summary>
        /// Only fpr Macintosh compatibility, used only during open, 
        /// when true dxaPgn and dyaPgn are valid page number locations
        /// </summary>
        public bool fAutoPgn;

        /// <summary>
        /// Page number format code:<br/>
        /// 0 Arabic<br/>
        /// 1 Roman (upper case)<br/>
        /// 2 Roman (lower case)<br/>
        /// 3 Letter (upper case)<br/>
        /// 4 Letter (lower case)
        /// </summary>
        public byte nfcPgn;

        /// <summary>
        /// Set to true, when a section in a locked document is unlocked
        /// </summary>
        public bool fUnlocked;

        /// <summary>
        /// Chapter number seperator for page numbers
        /// </summary>
        public byte cnsPgn;

        /// <summary>
        /// Set to true wehn page numbering should be restarted 
        /// a the beginning of this section
        /// </summary>
        public bool fPgnRestart;

        /// <summary>
        /// When true, footnotes placed at end of section.<br/>
        /// When false, footnotes are placed at bottom of page.
        /// </summary>
        public bool fEndnote;

        /// <summary>
        /// Line numbering code:<br/>
        /// 0 Per page<br/>
        /// 1 Restart<br/>
        /// 2 Continue
        /// </summary>
        public byte lnc;

        /// <summary>
        /// Specification of which headers and footers are included in this section.<br/>
        /// (No longer used)
        /// </summary>
        public byte grpfIhdtSepOld;

        /// <summary>
        /// If 0, no line numbering, otherwise this is the line number modulus
        /// </summary>
        public UInt16 nLnnMod;

        /// <summary>
        /// Distance of
        /// </summary>
        public Int32 dxaLnn;

        /// <summary>
        /// When fAutoPgn is true, gives the x position of auto page number on page in twips
        /// </summary>
        public Int16 dxaPgn;

        /// <summary>
        /// When fAutoPgn is true, gives the y position of auto page number on page in twips
        /// </summary>
        public Int16 dyaPgn;

        /// <summary>
        /// When true, draw vertical lines between columns
        /// </summary>
        public bool fLBetween;

        /// <summary>
        /// Bin number supplied from windows printer driver indicating 
        /// which bin the first page of section will be printed
        /// </summary>
        public UInt16 dmBinFirst;

        /// <summary>
        /// Bin number supplied from windows printer driver indicating
        /// which bin the pages other than the first page of section will be printed
        /// </summary>
        public UInt16 dmBinOther;

        /// <summary>
        /// dmPaper code for form selected by user
        /// </summary>
        public UInt16 dmPaperReq;

        /// <summary>
        /// When true, properties have been changed with revision marking on
        /// </summary>
        public bool fPropRMark;

        /// <summary>
        /// Index to author IDs stored in hsttbfRMark. used when properties have 
        /// been changed when revision marking was enabled
        /// </summary>
        public Int16 ibstPropRMark;

        /// <summary>
        /// date/time at which properties of this were changed for this run of text by the author
        /// </summary>
        public DateAndTime dttmPropRMark;

        /// <summary>
        /// How big is a character grid unit (East Asian)
        /// </summary>
        public Int32 dxtCharSpace;

        /// <summary>
        /// Line ptch: How tall a grid unit is up/down
        /// </summary>
        public Int32 dyaLinePitch;

        /// <summary>
        /// Grid description:<br/>
        /// 0 Default<br/>
        /// 1 Chars and line<br/>
        /// 2 Lines only<br/>
        /// 3 Enforce Grid<br/>
        /// </summary>
        public UInt16 clm;

        /// <summary>
        /// Orientation of pages in that section:<br/>
        /// 0 Portrait<br/>
        /// 1 Landscape
        /// </summary>
        public byte dmOrientPage;

        /// <summary>
        /// heading number level for page number
        /// </summary>
        public byte iHeadingPgn;

        /// <summary>
        /// User specified starting page number
        /// </summary>
        public UInt16 pgnStart;

        /// <summary>
        /// Beginning line number for section
        /// </summary>
        public Int16 lnnMin;

        /// <summary>
        /// Page border properties
        /// </summary>
        public Int16 pgbProp;

        /// <summary>
        /// Page border applies to:<br/>
        /// 0 all pages in this section<br/>
        /// 1 first page in this section<br/>
        /// 2 all pages in this section but first<br/>
        /// 3 whole document (all sections)
        /// </summary>
        public Int16 pgbApplayTo;

        /// <summary>
        /// Page border depth:<br/>
        /// 0 in front<br/>
        /// 1 in back
        /// </summary>
        public Int16 pgbPageDepth;

        /// <summary>
        /// Page border offset from:<br/>
        /// 0 offset from text<br/>
        /// 1 offset from edge of page
        /// </summary>
        public Int16 pgbOffsetFrom;

        /// <summary>
        /// Width of page
        /// </summary>
        public UInt32 xaPage;

        /// <summary>
        /// Height of page
        /// </summary>
        public UInt32 yaPage;

        /// <summary>
        /// Used internally by Word
        /// </summary>
        public UInt32 xaPageNUp;

        /// <summary>
        /// Used internally by Word
        /// </summary>
        public UInt32 yaPageNUp;

        /// <summary>
        /// Text flow:<br/>
        /// 0 Horizontal with no @ fonts<br/>
        /// 1 Top to bottom with @ fonts<br/>
        /// 2 Bottom to top with no @ fonts<br/>
        /// 3 Top to bottom with no @ fonts<br/>
        /// 4 Horizontal with @ fonts<br/>
        /// 5 Vertical with no @ fonts
        /// </summary>
        public UInt16 wTextFlow;

        /// <summary>
        /// Left margin
        /// </summary>
        public UInt32 dxaLeft;

        /// <summary>
        /// Right margin
        /// </summary>
        public UInt32 dxaRight;

        /// <summary>
        /// Top margin
        /// </summary>
        public Int32 dyaTop;

        /// <summary>
        /// Bottom margin
        /// </summary>
        public Int32 dyaBottom;

        /// <summary>
        /// Gutter width
        /// </summary>
        public UInt32 dzaGutter;

        /// <summary>
        /// Y position of top header measured from top edge of page
        /// </summary>
        public UInt32 dyaHdrTop;

        /// <summary>
        /// Y position of bottom header measured from top edge of page
        /// </summary>
        public UInt32 dyaHdrBottom;

        /// <summary>
        /// Number of columns in section -1
        /// </summary>
        public Int16 ccolM1;

        /// <summary>
        /// When true, columns are evenly spaced
        /// </summary>
        public bool fEvenlySpaced;

        /// <summary>
        /// Vertical justification code:<br/>
        /// 0 top justified<br/>
        /// 1 centered<br/>
        /// 2 fully justified vertically<br/>
        /// 3 bottom justified
        /// </summary>
        public byte vjc;

        /// <summary>
        /// Distance that will be maintained between columns
        /// </summary>
        public Int32 dxaColumns;

        /// <summary>
        /// When true, section direction is right-to-left
        /// </summary>
        public bool fBiDi;

        /// <summary>
        /// When true, section has facing columns
        /// </summary>
        public bool fFacingCol;

        /// <summary>
        /// When true, section has a right-to-left gutter
        /// </summary>
        public bool fRTFGutter;

        /// <summary>
        /// When true, section has a right-to-left alignment
        /// </summary>
        public bool fRTFAlignment;

        /// <summary>
        /// Used internally by Word
        /// </summary>
        public Int32 dxaColumnWidth;

        /// <summary>
        /// Page orientation:<br/>
        /// 1 Portrait<br/>
        /// 2 Landscape<br/>
        /// 3 Mixed
        /// </summary>
        public byte dmOrientFirst;

        /// <summary>
        /// Top page border
        /// </summary>
        public BorderCode brcTop;

        /// <summary>
        /// Left page border
        /// </summary>
        public BorderCode brcLeft;

        /// <summary>
        /// Bottom page border
        /// </summary>
        public BorderCode brcBottom;

        /// <summary>
        /// Right page border
        /// </summary>
        public BorderCode brcRight;
                    
        /// <summary>
        /// Multilevel auto numbering list data
        /// </summary>
        public OutlineLiSTData olstAnm;

        /// <summary>
        /// Used for section property revision marking. 
        /// The SEP at the time fHasOldProps is true, the is the old SEP.
        /// </summary>
        public bool fhasOldProps;

        /// <summary>
        /// Starting footnote number
        /// </summary>
        public UInt16 nFtn;

        /// <summary>
        /// Number format for footnote reference
        /// </summary>
        public Int16 nfcFtnRef;

        /// <summary>
        /// Starting endnote number
        /// </summary>
        public UInt16 nEdn;

        /// <summary>
        /// Number format for endnote reference
        /// </summary>
        public Int16 nfcEdnRef;

        /// <summary>
        /// Creates a new SectionProperties with default values
        /// </summary>
        public SectionProperties()
        {
            setDefaultValues();
        }

        private void setDefaultValues()
        {
            //The standard SEP is all zero ...
            this.brcBottom = new BorderCode();
            this.brcLeft = new BorderCode();
            this.brcRight = new BorderCode();
            this.brcTop = new BorderCode();
            this.ccolM1 = 0;
            this.clm = 0;
            this.cnsPgn = 0;
            this.dmBinFirst = 0;
            this.dmBinOther = 0;
            this.dmOrientFirst = 0;
            this.dmPaperReq = 0;
            this.dttmPropRMark = new DateAndTime();
            this.dxaColumnWidth = 0;
            this.dxaLnn = 0;
            this.dxtCharSpace = 0;
            this.dyaLinePitch = 0;
            this.dzaGutter = 0;
            this.fAutoPgn = false;
            this.fBiDi = false;
            this.fFacingCol = false;
            this.fhasOldProps = false;
            this.fLBetween = false;
            this.fPgnRestart = false;
            this.fPropRMark = false;
            this.fRTFAlignment = false;
            this.fRTFGutter = false;
            this.fTitlePage = false;
            this.fUnlocked = false;
            this.grpfIhdtSepOld = 0;
            this.ibstPropRMark = 0;
            this.iHeadingPgn = 0;
            this.lnc = 0;
            this.lnnMin = 0;
            this.nEdn = 0;
            this.nfcEdnRef = 0;
            this.nfcFtnRef = 0;
            this.nfcPgn = 0;
            this.nFtn = 0;
            this.nLnnMod = 0;
            this.olstAnm = new OutlineLiSTData();
            this.pgbApplayTo = 0;
            this.pgbOffsetFrom = 0;
            this.pgbPageDepth = 0;
            this.pgbProp = 0;
            this.vjc = 0;
            this.wTextFlow = 0;

            //except as follows ...
            this.bkc = 2;
            this.dyaPgn = 720;
            this.dxaPgn = 720;
            this.fEndnote = true;
            this.fEvenlySpaced = true;
            this.xaPage = 12240;
            this.yaPage = 15840;
            this.xaPageNUp = 12240;
            this.yaPageNUp = 15840;
            this.dyaHdrTop = 720;
            this.dyaHdrBottom = 720;
            this.dmOrientPage = 1;
            this.dxaColumns = 720;
            this.dyaTop = 1440;
            this.dxaLeft = 1800;
            this.dyaBottom = 1440;
            this.dxaRight = 1800;
            this.pgnStart = 1;
        }


    }
}
