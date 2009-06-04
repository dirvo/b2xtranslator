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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class PlotAreaMapping : AbstractOpenXmlMapping,
          IMapping<ChartFormatsSequence>
    {
        ExcelContext _xlsContext;
        ChartPart _chartPart;

        bool _isChartsheet;

        public PlotAreaMapping(ExcelContext xlsContext, ChartPart chartPart, bool isChartsheet)
            : base(chartPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartPart = chartPart;

            this._isChartsheet = isChartsheet;
        }

        #region IMapping<ChartFormatsSequence> Members

        /// <summary>
        /// <complexType name="CT_PlotArea">
        ///     <sequence>
        ///         <element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
        ///         <choice minOccurs="1" maxOccurs="unbounded">
        ///             <element name="areaChart" type="CT_AreaChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="area3DChart" type="CT_Area3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="lineChart" type="CT_LineChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="line3DChart" type="CT_Line3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="stockChart" type="CT_StockChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="radarChart" type="CT_RadarChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="scatterChart" type="CT_ScatterChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="pieChart" type="CT_PieChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="pie3DChart" type="CT_Pie3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="doughnutChart" type="CT_DoughnutChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="barChart" type="CT_BarChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="bar3DChart" type="CT_Bar3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="ofPieChart" type="CT_OfPieChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="surfaceChart" type="CT_SurfaceChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="surface3DChart" type="CT_Surface3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="bubbleChart" type="CT_BubbleChart" minOccurs="1" maxOccurs="1"/>
        ///         </choice>
        ///         <choice minOccurs="0" maxOccurs="unbounded">
        ///             <element name="valAx" type="CT_ValAx" minOccurs="1" maxOccurs="1"/>
        ///             <element name="catAx" type="CT_CatAx" minOccurs="1" maxOccurs="1"/>
        ///             <element name="dateAx" type="CT_DateAx" minOccurs="1" maxOccurs="1"/>
        ///             <element name="serAx" type="CT_SerAx" minOccurs="1" maxOccurs="1"/>
        ///         </choice>
        ///         <element name="dTable" type="CT_DTable" minOccurs="0" maxOccurs="1"/>
        ///         <element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
        ///         <element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
        ///     </sequence>
        /// </complexType>
        /// </summary>
        public void Apply(ChartFormatsSequence chartFormatsSequence)
        {
            // c:plotArea
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElPlotArea, Dml.Chart.Ns);
            {
                // c:layout
                if (chartFormatsSequence.ShtProps.fManPlotArea && chartFormatsSequence.CrtLayout12A != null)
                {
                    chartFormatsSequence.CrtLayout12A.Convert(new LayoutMapping(this._xlsContext, this._chartPart, this._isChartsheet));
                }

                // chart groups
                foreach (AxisParentSequence axisParentSequence in chartFormatsSequence.AxisParentSequences)
                {
                    foreach (CrtSequence crtSequence in axisParentSequence.CrtSequences)
                    {
                        // The Chart3d record specifies that the plot area, axis group, and chart group are rendered 
                        // in a 3-D scene, rather than a 2-D scene, and specifies properties of the 3-D scene. If this 
                        // record exists in the chart sheet substream, the chart sheet substream MUST have exactly one 
                        // chart group. This record MUST NOT exist in a bar of pie, bubble, doughnut,
                        // filled radar, pie of pie, radar, or scatter chart group.
                        //
                        bool is3DChart = (crtSequence.Chart3d != null);

                        // area chart
                        if (crtSequence.ChartType is Area)
                        {
                        }
                        // bar and column chart
                        else if (crtSequence.ChartType is Bar)
                        {
                            crtSequence.Convert(new BarChartMapping(this._xlsContext, this._chartPart, chartFormatsSequence, this._isChartsheet, is3DChart));
                        }
                        // OfPieChart (Bar of pie / Pie of Pie)
                        else if (crtSequence.ChartType is BopPop)
                        {
                        }
                        // bubbleChart
                        else if (crtSequence.ChartType is Scatter && ((Scatter)crtSequence.ChartType).fBubbles)
                        {
                        }
                        // scatterChart
                        else if (crtSequence.ChartType is Scatter && !((Scatter)crtSequence.ChartType).fBubbles)
                        {
                        }
                        // lineChart and stockChart
                        else if (crtSequence.ChartType is Line)
                        {
                        }
                        // doughnutChart
                        else if (crtSequence.ChartType is Pie && ((Pie)crtSequence.ChartType).pcDonut != 0)
                        {
                        }
                        // pieChart
                        else if (crtSequence.ChartType is Pie && ((Pie)crtSequence.ChartType).pcDonut == 0)
                        {
                        }
                        // radarChart
                        else if (crtSequence.ChartType is Radar || crtSequence.ChartType is RadarArea)
                        {
                            // RadarArea (or "Filled Radar") has the radarStyle set to "filled")
                        }
                        // surfaceChart
                        else if (crtSequence.ChartType is Surf)
                        {
                        }
                    }
                }

                // axis groups
                foreach (AxisParentSequence axisParentSequence in chartFormatsSequence.AxisParentSequences)
                {
                    // NOTE: AxisParent.iax must be 0 for the primary axis group
                    AxesSequence axesSequence = axisParentSequence.AxesSequence;

                    if (axesSequence.IvAxisSequence != null)
                    {
                        // c:catAx
                        _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElCatAx, Dml.Chart.Ns);
                        {
                            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElAxId, Dml.Chart.Ns);
                            _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, axesSequence.IvAxisSequence.Axis.AxisId.ToString());
                            _writer.WriteEndElement();


                            CatSerRange catSerRange = axesSequence.IvAxisSequence.CatSerRange;
                            // c:crossAx
                            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElCrossAx, Dml.Chart.Ns);
                            //_writer.WriteAttributeString(Dml.BaseTypes.AttrVal, catSerRange.);
                            _writer.WriteEndElement(); // c:crossAx

                        }
                        _writer.WriteEndElement(); // c:catAx
                    }

                    // c:valAx
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElValAx, Dml.Chart.Ns);
                    {
                        _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElAxId, Dml.Chart.Ns);
                        _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, axesSequence.DvAxisSequence.Axis.AxisId.ToString());
                        _writer.WriteEndElement();
                    }
                    _writer.WriteEndElement(); // c:valAx
                }
            }
            _writer.WriteEndElement(); // c:plotArea
        }
        #endregion
    }
}
