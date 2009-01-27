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
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using System.IO;
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public abstract class OpenXmlPartContainer
    {
        protected const string REL_PREFIX = "rId";
        protected const string EXT_PREFIX = "extId";
        protected const string REL_FOLDER = "_rels";
        protected const string REL_EXTENSION = ".rels";

        protected List<OpenXmlPart> _parts = new List<OpenXmlPart>();
        protected List<OpenXmlPart> _referencedParts = new List<OpenXmlPart>();
        protected List<ExternalRelationship> _externalRelationships = new List<ExternalRelationship>();
        protected static int _nextRelId = 1;

        protected OpenXmlPartContainer _parent = null;

        public virtual string TargetName 
        { 
            get 
            { 
                return ""; 
            } 
        }
        
        public virtual string TargetExt 
        {
            get
            { 
                return ""; 
            } 
        }
        
        public virtual string TargetDirectory 
        { 
            get 
            { 
                return ""; 
            }
            set
            {
            }
        }

        public virtual string TargetDirectoryAbsolute
        {
            get
            {
                // build complete path name from all parent parts
                string path = this.TargetDirectory;
                OpenXmlPartContainer part = this.Parent;
                while (part != null)
                {
                    path = Path.Combine(part.TargetDirectory, path);
                    part = part.Parent;
                }
                if (path == "ppt\\slides\\media") return "ppt\\media";
                if (path == "ppt\\slideMasters\\..\\slideLayouts") return "ppt\\slideLayouts";
                if (path == "ppt\\slideMasters\\..\\slideLayouts\\..\\media") return "ppt\\media";
                if (path == "ppt\\slides\\..\\media") return "ppt\\media";
                if (path == "ppt\\slideMasters\\..\\media") return "ppt\\media";
                return path;
            }
        }

        public virtual string TargetFullName
        {
            get
            {
                return Path.Combine(this.TargetDirectoryAbsolute, this.TargetName) + this.TargetExt;
            }
        }

        internal OpenXmlPartContainer Parent
        {
            get { return _parent; }
            set { _parent = value; }
        }

        protected IEnumerable<OpenXmlPart> Parts
        {
            get
            {
                return _parts;
            }
        }

        protected IEnumerable<OpenXmlPart> ReferencedParts
        {
            get
            {
                return _referencedParts;
            }
        }

        protected IEnumerable<ExternalRelationship> ExternalRelationships
        {
            get
            {
                return _externalRelationships;
            }
        }

        public virtual T AddPart<T>(T part) where T : OpenXmlPart
        {
            // generate a relId for the part 
            part.RelId = _nextRelId++;
            _parts.Add(part);

            if (part.HasDefaultContentType)
            {
                part.Package.AddContentTypeDefault(part.TargetExt.Replace(".", ""), part.ContentType);
            }
            else
            {
                part.Package.AddContentTypeOverride("/" + part.TargetFullName.Replace('\\', '/'), part.ContentType);
            }

            return part;
        }

        public ExternalRelationship AddExternalRelationship(string relationshipType, Uri externalUri)
        {
            ExternalRelationship rel = new ExternalRelationship(EXT_PREFIX + (_externalRelationships.Count + 1).ToString(), relationshipType, externalUri);
            _externalRelationships.Add(rel);
            return rel;
        }

        /// <summary>
        /// Add a part reference without actually managing the part.
        /// </summary>
        public virtual T ReferencePart<T>(T part) where T : OpenXmlPart
        {
            // We'll use the existing ID here.
            _referencedParts.Add(part);

            if (part.HasDefaultContentType)
            {
                part.Package.AddContentTypeDefault(part.TargetExt.Replace(".", ""), part.ContentType);
            }
            else
            {
                part.Package.AddContentTypeOverride("/" + part.TargetFullName.Replace('\\', '/'), part.ContentType);
            }

            return part;
        }

        protected virtual void WriteRelationshipPart(OpenXmlWriter writer)
        {
            List<OpenXmlPart> allParts = new List<OpenXmlPart>();
            allParts.AddRange(this.Parts);
            allParts.AddRange(this.ReferencedParts);

            // write part relationships
            if (allParts.Count > 0 || _externalRelationships.Count > 0)
            {
                string relFullName = Path.Combine(Path.Combine(this.TargetDirectoryAbsolute, REL_FOLDER), TargetName + TargetExt + REL_EXTENSION);
                writer.AddPart(relFullName);

                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", OpenXmlNamespaces.RelationsshipsPackage);

                foreach (ExternalRelationship rel in _externalRelationships)
                {
                    writer.WriteStartElement("Relationship", OpenXmlNamespaces.RelationsshipsPackage);
                    writer.WriteAttributeString("Id", rel.Id);
                    writer.WriteAttributeString("Type", rel.RelationshipType);

                    if (rel.Target.IsFile)
                    {
                        //reform the URI path for Word
                        //Word does not accept forward slahes in the path of a local file
                        writer.WriteAttributeString("Target", "file:///" + rel.Target.AbsolutePath.Replace("/", "\\"));
                    }
                    else
                    {
                        writer.WriteAttributeString("Target", rel.Target.ToString());
                    }

                    writer.WriteAttributeString("TargetMode", "External");

                    writer.WriteEndElement();
                }

                foreach (OpenXmlPart part in allParts)
                {
                    writer.WriteStartElement("Relationship", OpenXmlNamespaces.RelationsshipsPackage);
                    writer.WriteAttributeString("Id", part.RelIdToString);
                    writer.WriteAttributeString("Type", part.RelationshipType);

                    // write the target relative to the current part
                    writer.WriteAttributeString("Target", "/" + part.TargetFullName.Replace('\\', '/'));

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
    }
}
