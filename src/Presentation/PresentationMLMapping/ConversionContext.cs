using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.PptFileFormat;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class ConversionContext
    {
        private PresentationDocument _pptx;
        private XmlWriterSettings _writerSettings;
        private PowerpointDocument _ppt;

        /// <summary>
        /// The source of the conversion.
        /// </summary>
        public PowerpointDocument Ppt
        {
            get { return _ppt; }
            set { _ppt = value; }
        }

        /// <summary>
        /// This is the target of the conversion.<br/>
        /// The result will be written to the parts of this document.
        /// </summary>
        public PresentationDocument Pptx
        {
            get { return _pptx; }
            set { _pptx = value; }
        }

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return _writerSettings; }
            set { _writerSettings = value; }
        }

        public ConversionContext(PowerpointDocument ppt)
        {
            this.Ppt = ppt;
        }
    }
}
