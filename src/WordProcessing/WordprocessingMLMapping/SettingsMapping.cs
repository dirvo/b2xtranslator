using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class SettingsMapping : PropertiesMapping,
          IMapping<DocumentProperties>
    {
        private enum FootnotePosition
        {
            beneathText,
            docEnd,
            pageBottom,
            sectEnd
        }

        private enum ZoomType
        {
            none,
            fullPage,
            bestFit
        }

        private FileInformationBlock _fib;

        public SettingsMapping(ConversionContext ctx)
            : base(XmlWriter.Create(ctx.Docx.MainDocumentPart.AddSettingsPart().GetStream(), ctx.WriterSettings))
        {
            _fib = ctx.Doc.FIB;
        }

        public void Apply(DocumentProperties dop)
        {
            //start w:settings
            _writer.WriteStartElement("w", "settings", OpenXmlNamespaces.WordprocessingML);
            
            //zoom
            _writer.WriteStartElement("w", "zoom", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "percent", OpenXmlNamespaces.WordprocessingML, dop.wScaleSaved.ToString());
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ((ZoomType)dop.zkSaved).ToString());
            _writer.WriteEndElement();

            //default tab stop
            _writer.WriteStartElement("w", "defaultTabStop", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, dop.dxaTab.ToString());
            _writer.WriteEndElement();

            //proof state
            XmlElement proofState = _nodeFactory.CreateElement("w", "proofState", OpenXmlNamespaces.WordprocessingML);
            if (dop.fGramAllClean)
                appendValueAttribute(proofState, "grammar", "clean");
            if (proofState.Attributes.Count > 0)
                proofState.WriteTo(_writer);

            //footnote properties
            XmlElement footnotePr = _nodeFactory.CreateElement("w", "footnotePr", OpenXmlNamespaces.WordprocessingML);
            if (dop.nFtn != 0)
                appendValueAttribute(footnotePr, "numStart", dop.nFtn.ToString());
            if (dop.rncFtn != 0)
                appendValueAttribute(footnotePr, "numRestart", dop.rncFtn.ToString());
            if (dop.Fpc != 0)
                appendValueAttribute(footnotePr, "pos", ((FootnotePosition)dop.Fpc).ToString());
            if (footnotePr.Attributes.Count > 0)
                footnotePr.WriteTo(_writer);

            //compatibility settings
            _writer.WriteStartElement("w", "compat", OpenXmlNamespaces.WordprocessingML);

            //some settings must always be written
            _writer.WriteElementString("w", "useNormalStyleForList", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "doNotUseIndentAsNumberingTabStop", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "useAltKinsokuLineBreakRules", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "allowSpaceOfSameStyleInTable", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "doNotSuppressIndentation", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "doNotAutofitConstrainedTables", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "autofitToFirstFixedWidthCell", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "displayHangulFixedWidth", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "splitPgBreakAndParaMark", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "doNotVertAlignCellWithSp", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "doNotBreakConstrainedForcedTable", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "doNotVertAlignInTxbx", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "useAnsiKerningPairs", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteElementString("w", "cachedColBalance", OpenXmlNamespaces.WordprocessingML, "");
            
            //others are saved in the file
            if(!dop.fDontAdjustLineHeightInTable)
                _writer.WriteElementString("w", "adjustLineHeightInTable", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fAlignTablesRowByRow)
                _writer.WriteElementString("w", "alignTablesRowByRow", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fApplyBreakingRules)
                _writer.WriteElementString("w", "applyBreakingRules", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fUseAutoSpaceForFullWidthAlpha)
                _writer.WriteElementString("w", "autoSpaceLikeWord95", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fDntBlnSbDbWid)
                _writer.WriteElementString("w", "balanceSingleByteDoubleByteWidth", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fConvMailMergeEsc)
                _writer.WriteElementString("w", "convMailMergeEsc", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontBreakWrappedTables)
                _writer.WriteElementString("w", "doNotBreakWrappedTables", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fExpShRtn)
                _writer.WriteElementString("w", "doNotExpandShiftReturn", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fLeaveBackslashAlone)
                _writer.WriteElementString("w", "doNotLeaveBackslashAlone", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontSnapToGridInCell)
                _writer.WriteElementString("w", "doNotSnapToGridInCell", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontUseAsianBreakRules)
                _writer.WriteElementString("w", "doNotUseEastAsianBreakRules", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontUseHTMLParagraphAutoSpacing)
                _writer.WriteElementString("w", "doNotUseHTMLParagraphAutoSpacing", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontWrapTextWithPunct)
                _writer.WriteElementString("w", "doNotWrapTextWithPunct", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fFtnLayoutLikeWW8)
                _writer.WriteElementString("w", "footnoteLayoutLikeWW8", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fForgetLastTabAlign)
                _writer.WriteElementString("w", "forgetLastTabAlignment", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fGrowAutofit)
                _writer.WriteElementString("w", "growAutofit", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fLayoutRawTableWidth)
                _writer.WriteElementString("w", "layoutRawTableWidth", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fLayoutTableRowsApart)
                _writer.WriteElementString("w", "layoutTableRowsApart", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fLineWrapLikeWord6)
                _writer.WriteElementString("w", "lineWrapLikeWord6", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fMWSmallCaps)
                _writer.WriteElementString("w", "mwSmallCaps", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fNoColumnBalance)
                _writer.WriteElementString("w", "noColumnBalance", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fNoLeading)
                _writer.WriteElementString("w", "noLeading", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fNoSpaceRaiseLower)
                _writer.WriteElementString("w", "noSpaceRaiseLower", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fNoTabForInd)
                _writer.WriteElementString("w", "noTabHangInd", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fPrintBodyBeforeHdr)
                _writer.WriteElementString("w", "printBodyTextBeforeHeader", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fMapPrintTextColor)
                _writer.WriteElementString("w", "printColBlack", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontAllowFieldEndSelect)
                _writer.WriteElementString("w", "selectFldWithFirstOrLastChar", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fSpLayoutLikeWW8)
                _writer.WriteElementString("w", "shapeLayoutLikeWW8", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fShowBreaksInFrames)
                _writer.WriteElementString("w", "showBreaksInFrames", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fMakeSpaceForUL)
                _writer.WriteElementString("w", "spaceForUL", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fTruncDxaExpand)
                _writer.WriteElementString("w", "spacingInWholePoints", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fSubOnSize)
                _writer.WriteElementString("w", "subFontBySize", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fSuppressSpbfAfterPageBreak)
                _writer.WriteElementString("w", "suppressSpBfAfterPgBrk", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fSuppressTopSpacing)
                _writer.WriteElementString("w", "suppressTopSpacing", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fSwapBordersFacingPgs)
                _writer.WriteElementString("w", "swapBordersFacingPages", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fDntULTrlSpc)
                _writer.WriteElementString("w", "ulTrailSpace", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fPrintMet)
                _writer.WriteElementString("w", "usePrinterMetrics", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fUseWord2002TableStyleRules)
                _writer.WriteElementString("w", "useWord2002TableStyleRules", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fUserWord97LineBreakingRules)
                _writer.WriteElementString("w", "useWord97LineBreakRules", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fWPJust)
                _writer.WriteElementString("w", "wpJustification", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fWPSpace)
                _writer.WriteElementString("w", "wpSpaceWidth", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fWrapTrailSpaces)
                _writer.WriteElementString("w", "wrapTrailSpaces", OpenXmlNamespaces.WordprocessingML, "");
            _writer.WriteEndElement();
            
            //activeWritingStyle
            //alignBordersAndEdges
            //alwaysMergeEmptyNamespace
            //alwaysShowPlaceholderText
            //attachedSchema
            //attachedTemplate
            //autoFormatOverride
            //autoHyphenation
            //bookFoldPrinting
            //bookFoldPrintingSheets
            //bookFoldRevPrinting
            //bordersDoNotSurroundFooter
            //bordersDoNotSurroundHeader
            //captions
            //characterSpacingControl
            //clickAndTypeStyle
            //clrSchemeMapping
            //consecutiveHyphenLimit
            //decimalSymbol
            //defaultTableStyle
            //displayBackgroundShape
            //displayHorizontalDrawingGridEvery
            //displayVerticalDrawingGridEvery
            //documentProtection
            //documentType
            //docVars
            //doNotAutoCompressPictures
            //doNotDemarcateInvalidXml
            //doNotDisplayPageBoundaries
            //doNotEmbedSmartTags
            //doNotHyphenateCaps
            //doNotIncludeSubdocsInStats
            //doNotShadeFormData
            //doNotTrackFormatting
            //doNotTrackMoves
            //doNotUseMarginsForDrawingGridOrigin
            //doNotValidateAgainstSchema
            //drawingGridHorizontalOrigin
            //drawingGridHorizontalSpacing
            //drawingGridVerticalOrigin
            //drawingGridVerticalSpacing
            //embedSystemFonts
            //embedTrueTypeFonts
            //endnotePr
            //evenAndOddHeaders
            //forceUpgrade
            //formsDesign
            //gutterAtTop
            //hdrShapeDefaults
            //hideGrammaticalErrors
            //hideSpellingErrors
            //hyphenationZone
            //ignoreMixedContent
            //linkStyles
            //listSeparator
            //mailMerge
            //mathPr
            //mirrorMargins
            //noLineBreaksAfter
            //noLineBreaksBefore
            //noPunctuationKerning
            //printFormsData
            //printFractionalCharacterWidth
            //printPostScriptOverText
            //printTwoOnOne
            //readModeInkLockDown
            //removeDateAndTime
            //removePersonalInformation
            //revisionView
            //rsids
            //saveFormsData
            //saveInvalidXml
            //savePreviewPicture
            //saveSubsetFonts
            //saveThroughXslt
            //saveXmlDataOnly
            //schemaLibrary
            //shapeDefaults
            //showEnvelope
            //showXMLTags
            //smartTagType
            //strictFirstAndLastChars
            //styleLockQFSet
            //styleLockTheme
            //stylePaneFormatFilter
            //stylePaneSortMethod
            //summaryLength
            //themeFontLang 
            //trackRevisions
            //uiCompat97To2003
            //updateFields
            //useXSLTWhenSaving
            //view
            //writeProtection

            //close w:settings
            _writer.WriteEndElement();

            _writer.Flush();
        }
    }
}
