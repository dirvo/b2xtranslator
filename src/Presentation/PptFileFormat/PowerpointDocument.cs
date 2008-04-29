/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Reflection;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class PowerpointDocument : IVisitable, IEnumerable<Record>
    {
        static PowerpointDocument() {
            Record.UpdateTypeToRecordClassMapping(Assembly.GetExecutingAssembly(), typeof(PowerpointDocument).Namespace);
        }

        /// <summary>
        /// The stream "PowerPoint Document"
        /// </summary>
        public VirtualStream PowerpointDocumentStream;

        public List<Record> RootRecords = new List<Record>();

        public PowerpointDocument(StructuredStorageFile file)
        {
            this.PowerpointDocumentStream = file.GetStream("PowerPoint Document");

            while (this.PowerpointDocumentStream.Position != this.PowerpointDocumentStream.Length)
            {
                this.RootRecords.Add(Record.readRecord(this.PowerpointDocumentStream));
            }
        }

        override public string ToString()
        {
            StringBuilder result = new StringBuilder(base.ToString());

            foreach (Record record in this.RootRecords)
            {
                result.AppendLine();
                result.AppendLine();
                result.Append("Root Record: ");
                result.Append(record.ToString());
            }

            return result.ToString();
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        #region IEnumerable<Record> Members

        public IEnumerator<Record> GetEnumerator()
        {
            foreach (Record rootRecord in this.RootRecords)
                foreach (Record record in rootRecord)
                    yield return record;
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            foreach (Record record in this)
                yield return record;
        }

        #endregion
    }
}
