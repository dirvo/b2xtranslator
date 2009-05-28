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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;


namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Simple struct to hold the biff record data 
    /// </summary>
    struct BiffHeader
    {
        public RecordType id;
        public UInt16 length;
    }

    /// <summary>
    /// Abstract class which implements some extractor properties and methods 
    /// </summary>
    public abstract class Extractor
    {
        public VirtualStreamReader StreamReader;   

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="sum">workbookstream </param>
        public Extractor(VirtualStreamReader reader)
        {
            this.StreamReader = reader;
            if (StreamReader == null)
            {
                throw new ExtractorException(ExtractorException.NULLPOINTEREXCEPTION);
            }
        }

        /// <summary>
        /// extracts the data from the given stream !!!
        /// </summary>
        public abstract void extractData(); 
    }
}
