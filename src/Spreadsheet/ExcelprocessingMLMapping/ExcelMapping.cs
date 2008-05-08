using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.ExcelprocessingMLMapping; 


namespace DIaLOGIKa.b2xtranslator.ExcelprocessingMLMapping
{
    public abstract class ExcelMapping :
        AbstractOpenXmlMapping,
        IMapping<XlsDocument>
    {
        protected XlsDocument xls;
        protected ExcelContext xlscon;


        public ExcelMapping(ExcelContext xlscon, OpenXmlPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), xlscon.WriterSettings))
        {
            this.xlscon = xlscon; 
        }

        public abstract void Apply(XlsDocument xls); 
        }

    
}
