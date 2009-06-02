/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class PageSetupMapping : AbstractOpenXmlMapping,
          IMapping<PageSetupSequence>
    {
        ExcelContext _xlsContext;
        ChartsheetPart _chartsheetPart;

        public PageSetupMapping(ExcelContext xlsContext, ChartsheetPart targetPart)
            : base(targetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartsheetPart = targetPart; 
        }

        #region IMapping<PageSetupSequence> Members

        public void Apply(PageSetupSequence pageSetupSequence)
        {
            // page margins
            _writer.WriteStartElement(Sml.Sheet.ElPageMargins, Sml.Ns);

            if (pageSetupSequence.LeftMargin != null)
            {
                _writer.WriteAttributeString(Sml.Sheet.AttrLeft, pageSetupSequence.LeftMargin.value.ToString(CultureInfo.InvariantCulture));
            }
            if (pageSetupSequence.RightMargin != null)
            {
                _writer.WriteAttributeString(Sml.Sheet.AttrRight, pageSetupSequence.RightMargin.value.ToString(CultureInfo.InvariantCulture));
            }
            if (pageSetupSequence.TopMargin != null)
            {
                _writer.WriteAttributeString(Sml.Sheet.AttrTop, pageSetupSequence.TopMargin.value.ToString(CultureInfo.InvariantCulture));
            }
            if (pageSetupSequence.BottomMargin != null)
            {
                _writer.WriteAttributeString(Sml.Sheet.AttrBottom, pageSetupSequence.BottomMargin.value.ToString(CultureInfo.InvariantCulture));
            }
            if (pageSetupSequence.Setup != null)
            {
                _writer.WriteAttributeString(Sml.Sheet.AttrHeader, pageSetupSequence.Setup.numHdr.ToString(CultureInfo.InvariantCulture));
                _writer.WriteAttributeString(Sml.Sheet.AttrFooter, pageSetupSequence.Setup.numFtr.ToString(CultureInfo.InvariantCulture));
            }
            _writer.WriteEndElement();


            // page setup
            if (pageSetupSequence.Setup != null)
            {
                _writer.WriteStartElement(Sml.Sheet.ElPageSetup, Sml.Ns);

                if (pageSetupSequence.Setup.fUsePage)
                {
                    _writer.WriteAttributeString(Sml.Sheet.AttrFirstPageNumber, pageSetupSequence.Setup.iPageStart.ToString(CultureInfo.InvariantCulture));
                }

                _writer.WriteAttributeString(Sml.Sheet.AttrPaperSize, pageSetupSequence.Setup.iPaperSize.ToString(CultureInfo.InvariantCulture));

                if (!(pageSetupSequence.Setup.fNoPls || pageSetupSequence.Setup.fNoOrient))
                {
                    // If fNoPls is 1, the value is undefined and MUST be ignored. If fNoOrient is 1, the value is undefined and MUST be ignored.
                    _writer.WriteAttributeString(Sml.Sheet.AttrOrientation, pageSetupSequence.Setup.fPortrait ? "landscape" : "portrait");
                }
                _writer.WriteAttributeString(Sml.Sheet.AttrUseFirstPageNumber, pageSetupSequence.Setup.fUsePage ? "1" : "0");

                if (!pageSetupSequence.Setup.fNoPls)
                {
                    _writer.WriteAttributeString(Sml.Sheet.AttrHorizontalDpi, pageSetupSequence.Setup.iRes.ToString(CultureInfo.InvariantCulture));
                    _writer.WriteAttributeString(Sml.Sheet.AttrVerticalDpi, pageSetupSequence.Setup.iVRes.ToString(CultureInfo.InvariantCulture));
                }

                _writer.WriteEndElement();
            }

        }

        #endregion
    }
}
