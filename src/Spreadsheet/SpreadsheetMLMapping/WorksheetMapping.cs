/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class WorksheetMapping : AbstractOpenXmlMapping,
          IMapping<BoundSheetData>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public WorksheetMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddWorksheetPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Worksheet xml document 
        /// </summary>
        /// <param name="bsd">BoundSheetData</param>
        public void Apply(BoundSheetData bsd)
        {
            xlsContext.CurrentSheet = bsd; 
            _writer.WriteStartDocument();
            _writer.WriteStartElement("worksheet", OpenXmlNamespaces.WorkBookML);


            // default info 
            if (bsd.defaultColWidth >= 0 || bsd.defaultRowHeight >= 0)
            {
                 _writer.WriteStartElement("sheetFormatPr");

                if (bsd.defaultColWidth >= 0)
                {
                    double colWidht = (double)bsd.defaultColWidth;
                    _writer.WriteAttributeString("defaultColWidth", Convert.ToString(colWidht, CultureInfo.GetCultureInfo("en-US")));
                }
                if (bsd.defaultRowHeight >= 0)
                {
                    TwipsValue tv = new TwipsValue(bsd.defaultRowHeight); 
                    _writer.WriteAttributeString("defaultRowHeight", Convert.ToString(tv.ToPoints(), CultureInfo.GetCultureInfo("en-US")));
                }
                if (bsd.zeroHeight)
                {
                    _writer.WriteAttributeString("zeroHeight", "1"); 
                }
                if (bsd.customHeight)
                {
                    _writer.WriteAttributeString("customHeight", "1");
                }
                if (bsd.thickTop)
                {
                    _writer.WriteAttributeString("thickTop", "1");
                }
                if (bsd.thickBottom)
                {
                    _writer.WriteAttributeString("thickBottom", "1");
                }

                 _writer.WriteEndElement();
            }



            // Col info 
            if (bsd.colInfoDataTable.Count > 0)
            {
                _writer.WriteStartElement("cols");
                foreach (ColumnInfoData col in bsd.colInfoDataTable)
                {
                    _writer.WriteStartElement("col");
                    // write min and max 
                    // booth values are 0 based in the binary format and 1 based in the oxml format 
                    // so you have to add 1 to the value!

                        _writer.WriteAttributeString("min", (col.min + 1).ToString());
                        _writer.WriteAttributeString("max", (col.max + 1).ToString());

                        if (col.widht != 0)
                        {
                            double colWidht = (double)col.widht / 256; 
                            _writer.WriteAttributeString("width", Convert.ToString(colWidht, CultureInfo.GetCultureInfo("en-US")));
                        }
                    if (col.hidden)
                        _writer.WriteAttributeString("hidden", "1");

                    if (col.outlineLevel > 0)
                        _writer.WriteAttributeString("outlineLevel", col.outlineLevel.ToString());
                     
                    if (col.customWidth)
                        _writer.WriteAttributeString("customWidth", "1");
                     

                    if (col.bestFit)
                        _writer.WriteAttributeString("bestFit", "1");

                    if (col.phonetic)
                        _writer.WriteAttributeString("phonetic", "1");

                    if (col.style > 15)
                    {
                        
                        _writer.WriteAttributeString("style", Convert.ToString(col.style - this.xlsContext.XlsDoc.workBookData.styleData.XFCellStyleDataList.Count, CultureInfo.GetCultureInfo("en-US")));
                    }

                    _writer.WriteEndElement();
                }


                _writer.WriteEndElement();
            }


            // End col info 



            _writer.WriteStartElement("sheetData");
           //  bsd.rowDataTable.Values
            foreach (RowData row in bsd.rowDataTable.Values)
            {

                // write row start tag
                // Row 
                _writer.WriteStartElement("row");
                // the rowindex from the binary format is zero based, the ooxml format is one based 
                _writer.WriteAttributeString("r", (row.Row +1).ToString());
                if (row.height != null)
                {
                    if (row.customHeight)
                    {
                        _writer.WriteAttributeString("ht", Convert.ToString(row.height.ToPoints(), CultureInfo.GetCultureInfo("en-US")));
                        _writer.WriteAttributeString("customHeight", "1");
                    }

                }
                
                if (row.hidden)
                {
                    _writer.WriteAttributeString("hidden", "1"); 
                }
                if (row.outlineLevel > 0)
                {
                    _writer.WriteAttributeString("outlineLevel", row.outlineLevel.ToString());
                }
                if (row.collapsed)
                {
                    _writer.WriteAttributeString("collapsed", "1");
                }
                if (row.customFormat)
                {
                    _writer.WriteAttributeString("customFormat", "1");
                    if (row.style > 15)
                    {
                        _writer.WriteAttributeString("s", (row.style - this.xlsContext.XlsDoc.workBookData.styleData.XFCellStyleDataList.Count).ToString());
                    }
                }
                if (row.thickBot)
                {
                    _writer.WriteAttributeString("thickBot", "1");
                }
                if (row.thickTop)
                {
                    _writer.WriteAttributeString("thickTop", "1");
                }
                if (row.minSpan + 1 > 0 && row.maxSpan > 0 && row.minSpan +1 < row.maxSpan)
                {
                    _writer.WriteAttributeString("spans", (row.minSpan+1).ToString() + ":" + row.maxSpan.ToString());
                }

                row.Cells.Sort(); 
                foreach (AbstractCellData cell in row.Cells)
                {
                    // Col 
                    _writer.WriteStartElement("c");
                    _writer.WriteAttributeString("r", ExcelHelperClass.intToABCString((int)cell.Col, (cell.Row + 1).ToString()));

                    if (cell.TemplateID > 15)
                    {
                        _writer.WriteAttributeString("s", (cell.TemplateID - this.xlsContext.XlsDoc.workBookData.styleData.XFCellStyleDataList.Count).ToString());
                    }

                    if (cell is StringCell)
                    {
                        _writer.WriteAttributeString("t", "s");
                    }
                    if (cell is FormulaCell)
                    {
                        FormulaCell fcell = (FormulaCell) cell;


                        if (((FormulaCell)cell).calculatedValue is String)
                        {
                            _writer.WriteAttributeString("t", "str");
                        }
                        else if (((FormulaCell)cell).calculatedValue is double)
                        {
                            _writer.WriteAttributeString("t", "n");
                        }
                        else if (((FormulaCell)cell).calculatedValue is byte)
                        {
                            _writer.WriteAttributeString("t", "b");
                        }
                        else if (((FormulaCell)cell).calculatedValue is int)
                        {
                            _writer.WriteAttributeString("t", "e");
                        }


                        // <f>1</f> 
                        _writer.WriteStartElement("f");
                        if (!fcell.isSharedFormula)
                        {
                            String value = FormulaInfixMapping.mapFormula(fcell.PtgStack, this.xlsContext);


                            if (fcell.usesArrayRecord)
                            {
                                _writer.WriteAttributeString("t", "array");
                                _writer.WriteAttributeString("ref", ExcelHelperClass.intToABCString((int)cell.Col, (cell.Row + 1).ToString()));
                            }
                            if (value.Equals(""))
                            {
                                TraceLogger.Debug("Formula Parse Error in Row {0}\t Column {1}\t", cell.Row.ToString(), cell.Col.ToString());
                            }

                            _writer.WriteString(value);
                        }
                        /// If this cell is part of a shared formula 
                        /// 
                        else
                        {
                            SharedFormulaData sfd = bsd.checkFormulaIsInShared(cell.Row, cell.Col);
                            if (sfd != null)
                            {
                                // t="shared" 
                                _writer.WriteAttributeString("t", "shared");
                                //  <f t="shared" ref="C4:C11" si="0">H4+I4-J4</f> 
                                _writer.WriteAttributeString("si", sfd.ID.ToString());
                                if (sfd.RefCount == 0)
                                {
                                    /// Write value and reference 
                                    _writer.WriteAttributeString("ref", sfd.getOXMLFormatedData());

                                    String value = FormulaInfixMapping.mapFormula(sfd.PtgStack, this.xlsContext,sfd.rwFirst,sfd.colFirst);
                                    _writer.WriteString(value);

                                    sfd.RefCount++; 
                                }
                                
                            }
                            else
                            {
                                TraceLogger.Debug("Formula Parse Error in Row {0}\t Column {1}\t", cell.Row.ToString(), cell.Col.ToString());                          
                            }
                        }

                        _writer.WriteEndElement(); 
                        /// write down calculated value from a formula
                        /// 

                        _writer.WriteStartElement("v");

                        if (((FormulaCell)cell).calculatedValue is int)
                        {
                            _writer.WriteString(FormulaInfixMapping.getErrorStringfromCode((int)((FormulaCell)cell).calculatedValue)); 
                        }
                        else
                        {
                            _writer.WriteString(Convert.ToString(((FormulaCell)cell).calculatedValue, CultureInfo.GetCultureInfo("en-US")));
                        }

                        _writer.WriteEndElement(); 



                    }
                    else
                    {// Data !!! 
                        _writer.WriteElementString("v", cell.getValue());
                    }
                    // add a type to the c element if the formula returns following types 

                    _writer.WriteEndElement();  // close cell (c)  
                }


                _writer.WriteEndElement();  // close row 
            }          

            // close tags 
            _writer.WriteEndElement();      // close sheetData 


            // Add the mergecell part 
            //
            // - <mergeCells count="2">
            //        <mergeCell ref="B3:C3" /> 
            //        <mergeCell ref="E3:F4" /> 
            //     </mergeCells>
            if (bsd.MERGECELLSData != null)
            {
                _writer.WriteStartElement("mergeCells");
                _writer.WriteAttributeString("count", bsd.MERGECELLSData.cmcs.ToString());
                foreach (MergeCellData mcell in bsd.MERGECELLSData.mergeCellDataList)
                {
                    _writer.WriteStartElement("mergeCell");
                    _writer.WriteAttributeString("ref", mcell.getOXMLFormatedData());
                    _writer.WriteEndElement();
                }
                // close mergeCells Tag 
                _writer.WriteEndElement(); 
            }





            _writer.WriteEndElement();      // close worksheet
            _writer.WriteEndDocument();
            bsd.worksheetId = this.xlsContext.SpreadDoc.WorkbookPart.GetWorksheetPart().RelId;
            bsd.worksheetRef = this.xlsContext.SpreadDoc.WorkbookPart.GetWorksheetPart().RelIdToString;
            
            // close writer 
            _writer.Flush();
        }
    }
}
