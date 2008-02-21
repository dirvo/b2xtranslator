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

        // WordprocessingML content types
        public const string MainDocument = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        public const string Styles = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        public const string FontTable = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
        public const string WebSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
        public const string Settings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

        public const string CustomXmlProperties = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

        public const string Comments = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
  
        public const string Footnotes="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"; 
        public const string Endnotes = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

        public const string Header = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
        public const string Footer = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    }

    public static class OpenXmlNamespaces
    {
        // package namespaces
        public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        public const string RelationsshipsPackage = "http://schemas.openxmlformats.org/package/2006/relationships";

        public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // WordprocessingML namespaces
        public const string WordprocessingML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    }

    public static class OpenXmlRelationshipTypes
    {
        public const string CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

        public const string Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

        public const string MainDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
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
    }
}
