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

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public static class OpenXmlContentTypes
    {
        // default content types
        public const string Xml = "application/xml";

        // package content types
        public const string Relationships = "application/vnd.openxmlformats-package.relationships+xml";

        public const string CoreProperties = "application/vnd.openxmlformats-package.core-properties+xml";
        

        // general office document content types
        public const string ExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        public const string Theme = "application/vnd.openxmlformats-officedocument.theme+xml";

        public const string CustomXmlProperties = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

    }
     
    public static class WordprocessingMLContentTypes
    {
        // WordprocessingML content types
        public const string MainDocument = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        public const string Styles = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        public const string FontTable = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
        public const string WebSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
        public const string Settings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

        public const string Comments = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
  
        public const string Footnotes="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"; 
        public const string Endnotes = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

        public const string Header = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
        public const string Footer = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    }

    public static class SpreadsheetMLContentTypes
    {
        // SpreadsheetML content types
        public const string Workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"; 
        public const string Styles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        public const string Worksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        public const string SharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        public const string Connections = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
    }

    public static class PresentationMLContentTypes
    {
        // PresentationML content types
        public const string Presentation = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
        public const string Slide = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        public const string SlideMaster = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        public const string NotesSlide = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
        public const string SlideLayout = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
        public const string TableStyles = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
        public const string ViewProps = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
        public const string PresProps = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
    }

    public static class OpenXmlNamespaces
    {
        // package namespaces
        public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        public const string RelationsshipsPackage = "http://schemas.openxmlformats.org/package/2006/relationships";

        public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // WordprocessingML namespaces
        public const string WordprocessingML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public const string VectorML = "urn:schemas-microsoft-com:vml";
    }

    public static class OpenXmlRelationshipTypes
    {
        public const string CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

        public const string Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

        public const string OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string Styles="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public const string FontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"; 
        public const string WebSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
        public const string Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

        public const string CustomXml = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        public const string CustomXmlProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";

        public const string Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        
        public const string Footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
        public const string Endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";

        public const string Header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        public const string Footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

        public const string Image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        // PresentationML
        public const string SlideLayout = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        public const string Slide = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        public const string SlideMaster = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    }
}
