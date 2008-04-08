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

        public SettingsMapping(SettingsPart settingsPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(settingsPart.GetStream(), xws))
        {
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
            //compat
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
