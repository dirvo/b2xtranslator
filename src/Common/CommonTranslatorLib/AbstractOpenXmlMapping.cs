using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.CommonTranslatorLib
{
    public class AbstractOpenXmlMapping
    {
        protected XmlWriter _writer;

        public AbstractOpenXmlMapping(XmlWriter writer)
        {
            _writer = writer;
        }
    }
}
