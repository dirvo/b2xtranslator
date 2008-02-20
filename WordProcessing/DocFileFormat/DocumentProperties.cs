using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class DocumentProperties
    {
        /// <summary>
        /// True when facing pages should be printed
        /// </summary>
        public bool ffacingPages;

        /// <summary>
        /// True when window control is in effect
        /// </summary>
        public bool fWindowControl;

        /// <summary>
        /// true when doc is a main doc for Print Merge Helper
        /// </summary>
        public bool fPMHMainDoc;

        /// <summary>
        /// Default line suppression storage:<br/>
        /// 0 form letter line supression<br/>
        /// 1 no line supression<br/>
        /// (no longer used)
        /// </summary>
        public Int16 grfSuppression;

        /// <summary>
        /// Footnote position code:<br/>
        /// 0 print as endnotes<br/>
        /// 1 print as bottom of page<br/>
        /// 2 print immediately beneath text
        /// </summary>
        public Int16 Fpc;

        /// <summary>
        /// No longer used
        /// </summary>
        public Int16 grpfIhdt;

        /// <summary>
        /// Restart index for footnotes:<br/>
        /// 0 don't restart note numbering<br/>
        /// 1 restart for each section<br/>
        /// 2 restart for each page
        /// </summary>
        public Int16 rncFtn;

        /// <summary>
        /// Initial footnote number for document
        /// </summary>
        public Int16 nFtn;

        /// <summary>
        /// When true, indicates that information in the hplcpad should 
        /// be refreshed since outline has been dirtied
        /// </summary>
        public bool fOutlineDirtySave;

        /// <summary>
        /// When true, Word believes all pictures recorded in the 
        /// document were created on a Macintosh
        /// </summary>
        public bool fOnlyMacPics;

        /// <summary>
        /// When true, Word believes all pictures recorded in the 
        /// document were created in Windows
        /// </summary>
        public bool fOnlyWinPics;

        /// <summary>
        /// When true, document was created as a print 
        /// merge labels document
        /// </summary>
        public bool fLabelDoc;

        /// <summary>
        /// When true, Word is allowed to hyphenate words that are capitalized
        /// </summary>
        public bool fHyphCapitals;

        /// <summary>
        /// When true, Word will hyphenate newly typed 
        /// text as a background task
        /// </summary>
        public bool fAutoHyphen;

        /// <summary>
        /// 
        /// </summary>
        public bool fFormNoFields;

        /// <summary>
        /// When true, Word will merge styles from its template
        /// </summary>
        public bool fLinkStyles;

        /// <summary>
        /// Whent true, Word will mark revisions as the document is edited
        /// </summary>
        public bool fRevMarking;

        /// <summary>
        /// When true, always make backup when document saved
        /// </summary>
        public bool fBackup;

        /// <summary>
        /// When true, the results of the last Word Count execution are still exactly correct
        /// </summary>
        public bool fExactWords;

        /// <summary>
        /// When true, hidden documents contents are displayed
        /// </summary>
        public bool fPagHidden;

        /// <summary>
        /// When true, field results are displayed, when false, field codes are displayed
        /// </summary>
        public bool fPagResults;

        /// <summary>
        /// When true, annotations are locked for editing
        /// </summary>
        public bool fLockAtn;

        /// <summary>
        /// When true, swap margins on left/right pages
        /// </summary>
        public bool fMirrorMargins;

        /// <summary>
        /// When true, use TrueType fonts by default<br/>
        /// (flag obeyed only when doc was created by WinWord 2.x)
        /// </summary>
        public bool fDflttrueType;

        /// <summary>
        /// When true, file created with SUPPRESSTOPSPACING=YES in Win.ini<br/>
        /// (flag obeyed only when doc was created by WinWord 2.x)
        /// </summary>
        public bool fPagSuppressTopSpacing;

        /// <summary>
        /// When true, document is protected from edit operations
        /// </summary>
        public bool fProtEnabled;

        /// <summary>
        /// When true, restrict selections to occur only within form fields
        /// </summary>
        public bool fDispFormFldSel;

        /// <summary>
        /// When true, show revision markings on screen
        /// </summary>
        public bool fRMView;

        /// <summary>
        /// When true, show revision markings when document is printed
        /// </summary>
        public bool fRMPrint;

        /// <summary>
        /// When true, the current revision marking state is locked
        /// </summary>
        public bool fLockRev;

        /// <summary>
        /// When true, document contains embedded TrueType fonts
        /// </summary>
        public bool fEmbedFonts;

        /// <summary>
        /// Compatibility option: when true, don't add automatic tab 
        /// stops for hanging indent
        /// </summary>
        public bool fNoTabForInd;

        /// <summary>
        /// Compatibility option: when true, don't add extra space 
        /// for raised or lowered characters
        /// </summary>
        public bool fNoSpaceRaiseLower;

        /// <summary>
        /// Compatibility option: when true, suppress the paragraph 
        /// Space Before and Space After options after a page break
        /// </summary>
        public bool fSuppressSpbfAfterPageBreak;

        /// <summary>
        /// Compatibility option: when true, wrap trailing spaces 
        /// at the end of a line to the next line
        /// </summary>
        public bool fWrapTrailSpaces;

        /// <summary>
        /// Compatibility option: when true, print colors as black 
        /// on non-color printer
        /// </summary>
        public bool fMapPrintTextColor;

        /// <summary>
        /// Compatibility option: when true, don't balance columns 
        /// for Continuous Section starts
        /// </summary>
        public bool fNoColumnBalance;

        /// <summary>
        /// 
        /// </summary>
        public bool fConvMailMergeEsc;

        /// <summary>
        /// Compatibility option: when true, suppress extra line 
        /// spacing at top of page
        /// </summary>
        public bool fSuppressTopSpacing;

        /// <summary>
        /// Compatibility option: when true, combine table borders 
        /// like Word 5.x for the Macintosh
        /// </summary>
        public bool fOrigWordTableRules;

        /// <summary>
        /// Compatibility option: when true, don't blank area 
        /// between metafile pictures
        /// </summary>
        public bool fTransparentMetafiles;

        /// <summary>
        /// Compatibility option: when true, show hard page or 
        /// column breaks in frames
        /// </summary>
        public bool fShowBreaksInFrames;

        /// <summary>
        /// Compatibility option: when true, swap left and right 
        /// pages on odd facing pages
        /// </summary>
        public bool fSwapBordersFacingPgs;

        /// <summary>
        /// Default tab width
        /// </summary>
        public UInt16 dxaTab;

        /// <summary>
        /// Reserved
        /// </summary>
        public UInt16 wSpare;

        /// <summary>
        /// Width of hyphenation hot zone measured in twips
        /// </summary>
        public UInt16 dxaHotZ;

        /// <summary>
        /// Number of lines allowed to have consecutive hyphens
        /// </summary>
        public UInt16 cCOnsecHypLim;

        /// <summary>
        /// Reserved
        /// </summary>
        public UInt16 wSpare2;

        /// <summary>
        /// Date and time document was created
        /// </summary>
        public DateAndTime dttmCreated;

        /// <summary>
        /// Date and time document was last revised
        /// </summary>
        public DateAndTime dttmRevised;

        /// <summary>
        /// Date and time document was last printed
        /// </summary>
        public DateAndTime dttmLastPrint;

        /// <summary>
        /// Number of times document has ben revised since its creation
        /// </summary>
        public Int16 nRevision;

        /// <summary>
        /// Time document was last edited
        /// </summary>
        public Int32 tmEdited;

        /// <summary>
        /// Count of words tallied by last Word Count execution
        /// </summary>
        public Int32 cWords;

        /// <summary>
        /// Count of characters tallied by the last Word Count execution
        /// </summary>
        public Int32 cCh;

        /// <summary>
        /// Count of pages tallied by the last Word Count execution
        /// </summary>
        public Int16 cPg;

        /// <summary>
        /// Count of paragraphs tallied by the last Word Count execution
        /// </summary>
        public Int32 cParas;

        /// <summary>
        /// Restart endnote number code:<br/>
        /// 0 don't restart endnote numbering<br/>
        /// 1 restart for each section<br/>
        /// 2 restart for each page
        /// </summary>
        public Int16 rncEdn;

        /// <summary>
        /// Beginning endnote number
        /// </summary>
        public Int16 nEdn;

        /// <summary>
        /// Endnote position code:<br/>
        /// 0 display endnotes at end of section<br/>
        /// 3 display endnotes at the end of document
        /// </summary>
        public Int16 Epc;

        /// <summary>
        /// Number format code for auto footnotes.<br/>
        /// Use the Number Format Table.<br/>
        /// Note: Only the first 16 values in the table can be used.
        /// </summary>
        public Int16 nfcFtnRef;

        /// <summary>
        /// Number format code for auto endnotes.<br/>
        /// Use the Number Format Table.<br/>
        /// Note: Only the first 16 values in the table can be used.
        /// </summary>
        public Int16 nfcEdnRef;

        /// <summary>
        /// Only print data inside of form fields
        /// </summary>
        public bool fPrintFormData;

        /// <summary>
        /// Only save document data that is inside of a form field
        /// </summary>
        public bool fSaveFormData;

        /// <summary>
        /// Shade form fields
        /// </summary>
        public bool fShadwFormData;

        /// <summary>
        /// When true, include footnotes and endnotes in Word Count
        /// </summary>
        public bool fWCFtnEdn;

        /// <summary>
        /// Count of lines tallied by last Word Count operation
        /// </summary>
        public Int32 cLines;

        /// <summary>
        /// Count of words in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public Int32 cWordsFtnEdn;

        /// <summary>
        /// Count of characters in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public Int32 cChFtnEdn;

        /// <summary>
        /// Count of pages in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public Int32 cPgFtnEdn;

        /// <summary>
        /// Count of paragraphs in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public Int32 cParasFtnEdn;

        /// <summary>
        /// Count of lines in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public Int32 cLinesFtnEdn;

        /// <summary>
        /// Document protection password key only valid if 
        /// fProtEnabled, fLockAtn or fLockRev is true
        /// </summary>
        public Int32 lKeyProtDoc;

        /// <summary>
        /// Document view kind<br/>
        /// 0 Normal view<br/>
        /// 1 Outline view<br/>
        /// 2 Page view
        /// </summary>
        public Int16 wvSaved;

        /// <summary>
        /// Zoom percentage
        /// </summary>
        public Int16 wScaleSaved;

        /// <summary>
        /// Zoom type:<br/>
        /// 0 None<br/>
        /// 1 Full page<br/>
        /// 2 Page width
        /// </summary>
        public Int16 zkSaved;

        /// <summary>
        /// This is a vertical document<br/>
        /// (Word 6 and 96 only)
        /// </summary>
        public bool fRotateFontW6;

        /// <summary>
        /// Gutter position for this doc:<br/>
        /// 0 Side<br/>
        /// 1 Top
        /// </summary>
        public Int16 iGutterPos;

        /// <summary>
        /// SUpress extra line spacing at top of page like Word 5.x for the Macintosh
        /// </summary>
        public bool fSuppressTopSpacingMac5;

        /// <summary>
        /// Expand/Codense by whole number of points
        /// </summary>
        public bool fTruncDxaExpand;

        /// <summary>
        /// Print body text before header/footer
        /// </summary>
        public bool fPrintBodyBeforeHdr;

        /// <summary>
        /// Don't add leading (extra space) between rows of text
        /// </summary>
        public bool fNoLeading;

        /// <summary>
        /// USer larger small caps like Word 5.x for the Macintosh
        /// </summary>
        public bool fMWSmallCaps;
    }
}
