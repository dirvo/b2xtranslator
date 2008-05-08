using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat; 
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.ExcelprocessingMLMapping
{
    /// <summary>
    /// Includes some attributes and methods required by the mapping classes 
    /// </summary>
    public class ExcelContext
    {

        private XmlWriterSettings writerSettings;
        private XlsDocument xlsDoc;

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return writerSettings; }
            set { writerSettings = value; }
        }

        /// <summary>
        /// The XlsDocument 
        /// </summary>
        public XlsDocument XlsDoc
        {
            get { return xlsDoc; }
            set { this.xlsDoc = value; }
        }
    }


}
