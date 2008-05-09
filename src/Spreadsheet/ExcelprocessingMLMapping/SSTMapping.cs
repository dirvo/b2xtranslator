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
using DIaLOGIKa.b2xtranslator.ExcelprocessingMLMapping;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat; 
using System.Diagnostics;


namespace DIaLOGIKa.b2xtranslator.ExcelprocessingMLMapping
{

    public class SSTMapping : AbstractOpenXmlMapping,
          IMapping<WorkbookExtractor>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public SSTMapping(ExcelContext xlsContext)
            :base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddSharedStringPart().GetStream(), xlsContext.WriterSettings) )
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the sharedstring xml document 
        /// </summary>
        /// <param name="wbextr">Workbookextractor</param>
        public void Apply(WorkbookExtractor wbextr)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("sst");
            // count="x" uniqueCount="y" 
            _writer.WriteAttributeString("count",  wbextr.sst.cstTotal.ToString());
            _writer.WriteAttributeString("uniqueCount", wbextr.sst.cstUnique.ToString());



            // create the string entries 
            foreach (String var in wbextr.sst.StringList)
            {
                _writer.WriteStartElement("si" );
                _writer.WriteElementString("t", var);
                _writer.WriteEndElement();
               
            }            

            // close tags 
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            // close writer 
            _writer.Flush();
        }
    }

}
