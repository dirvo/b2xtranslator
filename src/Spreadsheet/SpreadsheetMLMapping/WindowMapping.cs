using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class WindowMapping : AbstractOpenXmlMapping,
          IMapping<WindowSequence>
    {
        ExcelContext _xlsContext;
        ChartsheetPart _chartsheetPart;
        int _window1Id = 0;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public WindowMapping(ExcelContext xlsContext, ChartsheetPart chartsheetPart, int window1Id)
            : base(chartsheetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartsheetPart = chartsheetPart;
            this._window1Id = window1Id;
        }

        #region IMapping<WindowSequence> Members

        public void Apply(WindowSequence windowSequence)
        {
            _writer.WriteStartElement(Sml.Sheet.ElSheetView, Sml.Ns);

            _writer.WriteAttributeString(Sml.Sheet.AttrTabSelected, windowSequence.Window2.fSelected ? "1" : "0");
            _writer.WriteAttributeString(Sml.Sheet.AttrWorkbookViewId, this._window1Id.ToString());
            // TODO: complete mapping

            _writer.WriteEndElement();
        }

        #endregion
    }
}
