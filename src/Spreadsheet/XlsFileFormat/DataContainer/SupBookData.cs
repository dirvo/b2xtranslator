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
    /// This class contains several information about the SUPBOOK BIFF Record 
    /// 
    /// </summary>
    public class SupBookData : IVisitable
    {
        private String virtPath;
        public String VirtPath
        {
            get { return this.virtPath; }
        }

        private String[] rgst;
        public String[] RGST
        {
            get { return this.rgst; }
        }

        private bool selfref;

        public bool SelfRef
        {
            get { return this.selfref; }
        }

        private LinkedList<XCTData> xctDataList;
        public LinkedList<XCTData> XCTDataList
        {
            get { return this.xctDataList; }
        }

        private LinkedList<String> externNames;
        public LinkedList<String> ExternNames
        {
            get { return this.externNames; }

        }

        public int ExternalLinkId;
        public String ExternalLinkRef;
        public int Number; 

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="supbook">SUPBOOK BIFF Record </param>
        public SupBookData(SUPBOOK supbook)
        {
            this.rgst = supbook.rgst;
            this.virtPath = supbook.virtpathstring;
            this.selfref = supbook.isselfreferencing;
            this.xctDataList = new LinkedList<XCTData>();
            this.externNames = new LinkedList<string>(); 
        }

        /// <summary>
        /// returns the value at the specified position
        /// </summary>
        /// <param name="index">searched index</param>
        /// <returns></returns>
        public String getRgstString(int index)
        {
            return this.rgst[index]; 
        }

        /// <summary>
        /// Add a XCT Data structure to the internal stack 
        /// </summary>
        /// <param name="xct"></param>
        public void addXCT(XCT xct)
        {
            XCTData xctdata = new XCTData(xct);
            this.xctDataList.AddLast(xctdata); 
        }

        public void addCRN(CRN crn)
        {
            this.xctDataList.Last.Value.addCRN(crn);           
        }

        public void addEXTERNNAME(EXTERNNAME extname)
        {
            this.externNames.AddLast(extname.extName); 
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SupBookData>)mapping).Apply(this);
        }

        #endregion


    }
}
