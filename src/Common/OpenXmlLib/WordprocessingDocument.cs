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
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingML
{
    public enum WordprocessingDocumentType
    {
        Document,
        MacroEnabledDocument,
        MacroEnabledTemplate,
        Template
    }

    public class WordprocessingDocument : OpenXmlPackage
    {
        protected WordprocessingDocumentType _documentType;
        protected CustomXmlPropertiesPart _customFilePropertiesPart;
        protected MainDocumentPart _mainDocumentPart;
        protected StyleDefinitionsPart _styleDefinitionsPart;

        protected WordprocessingDocument(string fileName)
            : base(fileName)
        {
            _mainDocumentPart = new MainDocumentPart(this);
            this.AddPart(_mainDocumentPart);

            _styleDefinitionsPart = new StyleDefinitionsPart(this);
            this.AddPart(_styleDefinitionsPart);
        }

        public static WordprocessingDocument Create(string fileName, WordprocessingDocumentType type)
        {
            WordprocessingDocument doc = new WordprocessingDocument(fileName);
            doc.DocumentType = type;

            return doc;
        }

        public WordprocessingDocumentType DocumentType
        {
            get { return _documentType; }
            set { _documentType = value; }
        }

        public CustomXmlPropertiesPart CustomFilePropertiesPart
        {
            get { return _customFilePropertiesPart; }
        }

        public MainDocumentPart MainDocumentPart
        {
            get { return _mainDocumentPart; }
        }

        public StyleDefinitionsPart StyleDefinitionsPart
        {
            get { return _styleDefinitionsPart; }
        }
    }
}
