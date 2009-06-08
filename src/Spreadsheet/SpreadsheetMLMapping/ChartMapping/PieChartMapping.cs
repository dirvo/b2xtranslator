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
using System.Globalization;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class PieChartMapping : AbstractChartGroupMapping
    {
        public PieChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Pie))
            {
                throw new Exception("Invalid chart type");
            }

            Pie pie = crtSequence.ChartType as Pie;

            // c:pieChart or c:pie3DChart 
            _writer.WriteStartElement(Dml.Chart.Prefix, this._is3DChart ? Dml.Chart.ElPie3DChart : Dml.Chart.ElPieChart, Dml.Chart.Ns);
            {
                // Axis Ids
                foreach (int axisId in crtSequence.ChartFormat.AxisIds)
                {
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElAxId, Dml.Chart.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, axisId.ToString());
                    _writer.WriteEndElement();
                }

                // varyColors
                writeValueElement(_writer, "varyColors", crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Pie Chart Series
                foreach (SeriesFormatSequence seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:dPt (Data Points)
                        for (int i = 1; i < seriesFormatSequence.SsSequence.Count; i++)
                        {
                            // write a dPt for each SsSequence
                            SsSequence sss = seriesFormatSequence.SsSequence[i];
                            sss.Convert(new DataPointMapping(this.WorkbookContext, this.ChartContext, i-1));
                        }

                        // c:val
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext));

                        // c:shape

                        _writer.WriteEndElement(); // c:ser
                    }
                }

                // firstSliceAng
                writeValueElement(_writer, "firstSliceAng", pie.anStart.ToString());
            }
            _writer.WriteEndElement();
        }
        #endregion
    }
}
