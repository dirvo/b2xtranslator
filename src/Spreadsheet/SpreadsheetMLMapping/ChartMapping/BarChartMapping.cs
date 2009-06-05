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

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class BarChartMapping : AbstractChartGroupMapping
    {
        public BarChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Bar))
            {
                throw new Exception("Invalid chart type");
            }

            Bar bar = crtSequence.ChartType as Bar;
            

            // c:barChart / c:bar3DChart
            _writer.WriteStartElement(Dml.Chart.Prefix, this._is3DChart ? Dml.Chart.ElBar3DChart : Dml.Chart.ElBarChart, Dml.Chart.Ns);
            {
                // EG_BarChartShared
                // c:barDir
                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElBarDir, Dml.Chart.Ns);
                _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, bar.fTranspose ? "bar" : "col");
                _writer.WriteEndElement(); // c:barDir

                // c:grouping TODO

                // c:varyColors
                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElVaryColors, Dml.Chart.Ns);
                _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, crtSequence.ChartFormat.fVaried ? "1" : "0");
                _writer.WriteEndElement(); // c:varyColors

                // Bar Chart Series
                foreach (SeriesFormatSequence seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:invertIfNegative (stored in AreaFormat)

                        // c:pictureOptions

                        // c:dPt (Data Point)

                        // c:dLbls (Data Labels)

                        // c:trendline 

                        // c:errBars

                        // c:cat (Category Axis Data)

                        // c:val
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext));

                        // c:shape

                        _writer.WriteEndElement(); // c:ser
                    }
                }


                // Data Labels


                if (this._is3DChart)
                {
                    // c:gapWidth
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElGapWidth, Dml.Chart.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, crtSequence.Chart3d.pcGap.ToString());
                    _writer.WriteEndElement(); // c:gapWidth

                    // c:gapDepth
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElGapDepth, Dml.Chart.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, crtSequence.Chart3d.pcDepth.ToString());
                    _writer.WriteEndElement(); // c:gapDepth

                    // c:shape (defined in Chart3DBarShape???)
                }
                else
                {
                    // c:gapWidth
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElGapWidth, Dml.Chart.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, bar.pcGap.ToString());
                    _writer.WriteEndElement(); // c:gapWidth

                    // c:overlap
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElOverlap, Dml.Chart.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, bar.pcOverlap.ToString());
                    _writer.WriteEndElement(); // c:overlap

                    // Series Lines


                }

                // Axis Ids
                foreach (int axisId in crtSequence.ChartFormat.AxisIds)
                {
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElAxId, Dml.Chart.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, axisId.ToString());
                    _writer.WriteEndElement();
                }
            }
            _writer.WriteEndElement();
        }
        #endregion
    }
}
