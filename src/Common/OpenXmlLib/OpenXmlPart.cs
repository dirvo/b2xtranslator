﻿/*
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
using System.IO;
using System.Xml;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public abstract class OpenXmlPart : OpenXmlPartContainer
    {
        protected int _relId = 0;
        protected MemoryStream _stream;
        protected XmlWriter _xmlWriter;

        public OpenXmlPart(OpenXmlPartContainer parent)
        {
            _parent = parent;
            _stream = new MemoryStream();

            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = false;
            xws.CloseOutput = false;
            xws.Encoding = Encoding.UTF8;
            xws.Indent = true;
            xws.ConformanceLevel = ConformanceLevel.Document;

            _xmlWriter = XmlWriter.Create(_stream, xws);
        }

        public override string TargetExt { get { return ".xml"; } }
        
        public abstract string ContentType { get; }
        public abstract string RelationshipType { get; }
        
        internal virtual bool HasDefaultContentType { get { return false; } }

        public Stream GetStream()
        {
            _stream.Seek(0, SeekOrigin.Begin);
            return _stream;
        }

        public XmlWriter XmlWriter
        {
            get
            {
                return _xmlWriter;
            }
        }

        public int RelId
        {
            get { return _relId; }
            set { _relId = value; }
        }

        public string RelIdToString
        {
            get { return REL_PREFIX + _relId.ToString(); }
        }

        public OpenXmlPackage Package
        {
            get
            {
                OpenXmlPartContainer partContainer = this.Parent;
                while (partContainer.Parent != null)
                {
                    partContainer = partContainer.Parent;
                }
                return partContainer as OpenXmlPackage;
            }
        }

        internal virtual void WritePart(OpenXmlWriter writer)
        {
            foreach (OpenXmlPart part in this.Parts)
            {
                part.WritePart(writer);
            }
            
            writer.AddPart(this.TargetFullName);

            writer.Write(this.GetStream());
            
            this.WriteRelationshipPart(writer);
        }
    }
}
