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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class WorkbookMapping : AbstractOpenXmlMapping,
          IMapping<WorkBookData>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public WorkbookMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Workbook xml document 
        /// </summary>
        /// <param name="bsd">WorkSheetData</param>
        public void Apply(WorkBookData bsd)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("workbook", OpenXmlNamespaces.WorkBookML);
            _writer.WriteAttributeString("xmlns", "r", "", OpenXmlNamespaces.Relationships); 
            _writer.WriteStartElement("sheets");

            foreach (WorkSheetData var in bsd.boundSheetDataList)
            {
           //     if (var.boundsheetRecord.sheetType == BOUNDSHEET.sheetTypes.worksheet)
             //   {
                    _writer.WriteStartElement("sheet");
                    _writer.WriteAttributeString("name", var.worksheetName);
                    _writer.WriteAttributeString("sheetId", var.worksheetId.ToString());

                    _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, var.worksheetRef);
                    _writer.WriteEndElement();
            //    }
            }
            _writer.WriteEndElement();      // close sheetData 

            bool ParentTagWritten = false; 
            if (bsd.supBookDataList.Count != 0)
            {
                
                /*
                    <externalReferences>
                        <externalReference r:id="rId4" /> 
                        <externalReference r:id="rId5" /> 
                    </externalReferences>
                 */
                foreach (SupBookData var in bsd.supBookDataList)
                {
                    if (!var.SelfRef)
                    {
                        if (!ParentTagWritten)
                        {
                            _writer.WriteStartElement("externalReferences");
                            ParentTagWritten = true; 
                        }
                        _writer.WriteStartElement("externalReference");
                        _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, var.ExternalLinkRef);
                        _writer.WriteEndElement();
                    }
                }
                if (ParentTagWritten)
                {
                    _writer.WriteEndElement();
                }
            }

            // write definedNames 

            if (bsd.definedNameList.Count > 0)
            {
                //<definedNames>
                //<definedName name="abc" comment="test" localSheetId="1">Sheet1!$B$3</definedName>
                //</definedNames>
                _writer.WriteStartElement("definedNames");

                foreach (DefinedNameData item in bsd.definedNameList)
                {
                    if (item.ptgStack.Count > 0)
                    {
                        _writer.WriteStartElement("definedName");
                        if (item.Name.Length > 1)
                        {
                            _writer.WriteAttributeString("name", item.Name);
                        }
                        else
                        {
                            string internName = "_xlnm." + ExcelHelperClass.getNameStringfromBuiltInFunctionID(item.Name);
                            _writer.WriteAttributeString("name", internName);
                        }
                        if (item.itab > 0)
                        {
                            _writer.WriteAttributeString("localSheetId", (item.itab - 1).ToString());
                        }
                        _writer.WriteValue(FormulaInfixMapping.mapFormula(item.ptgStack, xlsContext));

                        _writer.WriteEndElement();
                    }
                }


                _writer.WriteEndElement();     

            }



            // close tags 
           
            _writer.WriteEndElement();      // close worksheet
            _writer.WriteEndDocument();

            // close writer 
            _writer.Flush();
        }

    }
}
