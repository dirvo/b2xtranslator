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
using System.Collections;
using System.Text;
using System.IO;
using System.Collections.ObjectModel;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;


namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class is a container for the extracted data 
    /// Implements the IVisitable Interface 
    /// </summary>
    public class WorkBookData: IVisitable
    {
        /// <summary>
        /// This attribute stores the SharedStringTable Data
        /// </summary>
        private SSTData sstData;
        public SSTData SstData
        {
            get { return sstData; }
            set { sstData = value; }
        }

        public List<BoundSheetData> boundSheetDataList; 

        /// <summary>
        /// Ctor 
        /// </summary>
        public WorkBookData()
        {
            boundSheetDataList = new List<BoundSheetData>(); 
        }

        /// <summary>
        /// Adds a BoundSheetData Element to the internal list 
        /// </summary>
        /// <param name="bsd">The Boundsheetdata element</param>
        public void addBoundSheetData(BoundSheetData bsd)
        {
            this.boundSheetDataList.Add(bsd); 
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WorkBookData>)mapping).Apply(this);
        }

        #endregion
    }
}
