using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(1016)]
    public class MainMaster : Slide
    {
        public Dictionary<string, string> Layouts = new Dictionary<string, string>();
        public MainMaster(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {
                foreach (Record rec in Children)
                {
                    if (rec is RoundTripContentMasterInfo12)
                    {
                        RoundTripContentMasterInfo12 info = (RoundTripContentMasterInfo12)rec;
                        string xml = info.XmlDocumentElement.OuterXml;
                        xml = xml.Replace("http://schemas.openxmlformats.org/drawingml/2006/3/main", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        string title = info.XmlDocumentElement.Attributes["type"].InnerText;
                        Layouts.Add(title, xml);
                    }           
                }
        }
    }
}
