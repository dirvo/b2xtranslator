﻿/*
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

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class DrawingMapping : AbstractOpenXmlMapping,
          IMapping<ChartSheetContentSequence>
    {
        ExcelContext _xlsContext;
        DrawingsPart _drawingsPart;

        bool _isChartsheet;

        public DrawingMapping(ExcelContext xlsContext, DrawingsPart targetPart, bool isChartsheet)
            : base(targetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._drawingsPart = targetPart;

            this._isChartsheet = isChartsheet;
        }

        #region IMapping<ChartSheetContentSequence> Members

        public void Apply(ChartSheetContentSequence chartSheetContentSequence)
        {
            _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElWsDr, Dml.SpreadsheetDrawing.Ns);
            _writer.WriteAttributeString("xmlns", Dml.SpreadsheetDrawing.Prefix, "", Dml.SpreadsheetDrawing.Ns);
            _writer.WriteAttributeString("xmlns", Dml.Prefix, "", Dml.Ns);

            if (this._isChartsheet)
            {
                _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElAbsoluteAnchor, Dml.SpreadsheetDrawing.Ns);
                {
                    Chart chart = chartSheetContentSequence.ChartFormatsSequence.Chart;

                    // NOTE: Excel seems to somehow round the pos and ext values. The exact calculation is not documented.
                    //   Besides, Excel might write negative values which are corrected to 0 by Excel on load time.
                    //
                    // xdr:pos
                    _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElPos, Dml.SpreadsheetDrawing.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrX, Math.Max(0, new PtValue(chart.x.Value).ToEmu()).ToString());
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrY, Math.Max(0, new PtValue(chart.y.Value).ToEmu()).ToString());
                    _writer.WriteEndElement();

                    // xdr:ext
                    _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElExt, Dml.SpreadsheetDrawing.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrCx, Math.Max(0, new PtValue(chart.dx.Value).ToEmu()).ToString());
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrCy, Math.Max(0, new PtValue(chart.dy.Value).ToEmu()).ToString());
                    _writer.WriteEndElement();

                    _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElGraphicFrame, Dml.SpreadsheetDrawing.Ns);
                    {
                        // TODO: add graphic properties
                        _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElNvGraphicFramePr, Dml.SpreadsheetDrawing.Ns);
                        {
                            _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCNvPr, Dml.SpreadsheetDrawing.Ns);
                            _writer.WriteAttributeString(Dml.DocumentProperties.AttrId, this._drawingsPart.RelId.ToString());
                            _writer.WriteAttributeString(Dml.DocumentProperties.AttrName, "Shape");
                            _writer.WriteEndElement(); // xdr:cNvPr

                            _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCNvGraphicFramePr, Dml.SpreadsheetDrawing.Ns);
                            _writer.WriteStartElement(Dml.Prefix, Dml.DocumentProperties.ElGraphicFrameLocks, Dml.Ns);
                            _writer.WriteAttributeString(Dml.DocumentProperties.AttrNoGrp, "1");
                            _writer.WriteEndElement(); // a:graphicFrameLocks
                            _writer.WriteEndElement(); // xdr:cNvGraphicFramePr
                        }
                        _writer.WriteEndElement(); // xdr:nvGraphicFramePr

                        // xdr:xfrm
                        _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElXfrm, Dml.SpreadsheetDrawing.Ns);
                        {
                            _writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElOff, Dml.Ns);
                            _writer.WriteAttributeString(Dml.BaseTypes.AttrX, "0");
                            _writer.WriteAttributeString(Dml.BaseTypes.AttrY, "0");
                            _writer.WriteEndElement(); // a:off

                            _writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElExt, Dml.Ns);
                            _writer.WriteAttributeString(Dml.BaseTypes.AttrCx, "0");
                            _writer.WriteAttributeString(Dml.BaseTypes.AttrCy, "0");
                            _writer.WriteEndElement(); // a:ext
                        }
                        _writer.WriteEndElement(); // xdr:xfrm


                        _writer.WriteStartElement(Dml.GraphicalObject.ElGraphic, Dml.Ns);
                        {
                            _writer.WriteStartElement(Dml.GraphicalObject.ElGraphicData, Dml.Ns);
                            _writer.WriteAttributeString(Dml.GraphicalObject.AttrUri, Dml.Chart.Ns);

                            // create and convert chart part
                            ChartPart chartPart = this._drawingsPart.AddChartPart();
                            ChartContext chartContext = new ChartContext(chartPart, chartSheetContentSequence, this._isChartsheet ? ChartContext.ChartLocation.Chartsheet : ChartContext.ChartLocation.Embedded);
                            chartSheetContentSequence.Convert(new ChartMapping(this._xlsContext, chartContext));

                            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElChart, Dml.Chart.Ns);
                            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, chartPart.RelIdToString);

                            _writer.WriteEndElement(); // c:chart

                            _writer.WriteEndElement(); // a:graphicData
                        }
                        _writer.WriteEndElement(); // a:graphic
                    }
                    _writer.WriteEndElement(); // a:graphicFrame

                    _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElClientData, Dml.SpreadsheetDrawing.Ns, string.Empty);
                }
                _writer.WriteEndElement(); // absoluteAnchor
            }
            else
            {
                // embedded drawing
            }

            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        #endregion
    }
}
