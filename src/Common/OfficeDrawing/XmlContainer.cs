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
using System.IO;
using System.IO.Compression;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// Regular containers are containers with Record children.<br/>
    /// (There also is containers that only have a zipped XML payload.
    /// </summary>
    public class XmlContainer : Record
    {
        public XmlElement XmlDocumentElement;

        public XmlContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            // Note: XmlContainers contain the data of a partial "unfinished"
            // OOXML file (.zip based) as their body.
            //
            // I really don't like writing the data to a temp file just to
            // be able to open it via ZipUtils.
            //
            // Possible alternatives:
            // 1) Using System.IO.Compression -- supports inflation, but can't parse Zip header data
            // 2) Modifying zlib + minizlib + ZipLib so I can pass in bytes, possible, but not worth the effort            

            string tempPath = Path.GetTempFileName();

            try
            {
                using (BinaryWriter tempStream = new BinaryWriter(new FileStream(tempPath, FileMode.Create)))
                {
                    int count = (int)this.Reader.BaseStream.Length;
                    byte[] bytes = this.Reader.ReadBytes(count);

                    tempStream.Write(bytes);
                }

                ZipReader zipReader = ZipFactory.OpenArchive(tempPath);
                Stream relStream = zipReader.GetEntry("_rels/.rels");

                XmlDocument relDocument = new XmlDocument();
                relDocument.Load(relStream);

                XmlNodeList rels = relDocument["Relationships"].GetElementsByTagName("Relationship");

                if (rels.Count != 1)
                    throw new Exception("Expected actly one Relationship in XmlContainer OOXML doc");

                string partPath = rels[0].Attributes["Target"].Value;
                Stream partStream = zipReader.GetEntry(partPath);

                XmlDocument partDoc = new XmlDocument();
                partDoc.Load(partStream);

                this.XmlDocumentElement = partDoc.DocumentElement;
            }
            finally
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch (IOException) { /* OK */ }
            }
        }
    }
}