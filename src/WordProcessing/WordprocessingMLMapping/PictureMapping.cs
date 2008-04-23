using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.IO;
using System.Drawing;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class PictureMapping
        : AbstractOpenXmlMapping,
          IMapping<PictureDescriptor>
    {
        ImagePart _xmlPart;

        public PictureMapping(XmlWriter writer, ImagePart xmlPart)
            : base(writer)
        {
            _xmlPart = xmlPart;
        }

        public void Apply(PictureDescriptor pict)
        {
            //start w:pict
            _writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

            //start v:rect
            _writer.WriteStartElement("v", "rect", OpenXmlNamespaces.VectorML);

            //start style attribute
            _writer.WriteStartAttribute("style");
            StringBuilder style = new StringBuilder();

            //size
            style.Append("width:");
            double widthPt = new TwipsValue(pict.dxaGoal * 0.001 * pict.mx).ToPoints();
            style.Append(widthPt.ToString(new CultureInfo("en-US")));
            style.Append("pt;");
            style.Append("height:");
            double heightPt = new TwipsValue(pict.dyaGoal * 0.001 * pict.my).ToPoints();
            style.Append(heightPt.ToString(new CultureInfo("en-US")));
            style.Append("pt;");

            //end style
            _writer.WriteString(style.ToString());
            _writer.WriteEndAttribute();

            //write image
            _writer.WriteStartElement("v", "imagedata", OpenXmlNamespaces.VectorML);

            //id
            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, _xmlPart.RelIdToString);

            //cropping
            if (pict.dyaCropTop != 0)
                _writer.WriteAttributeString("croptop", pict.dyaCropTop.ToString());
            if (pict.dxaCropRight != 0)
                _writer.WriteAttributeString("cropright", pict.dxaCropRight.ToString());
            if (pict.dyaCropBottom != 0)
                _writer.WriteAttributeString("cropbottom", pict.dyaCropBottom.ToString());
            if (pict.dxaCropLeft != 0)
                _writer.WriteAttributeString("cropleft", pict.dxaCropLeft.ToString());

            _writer.WriteEndElement();

            //close v:rect
            _writer.WriteEndElement();

            //close w:pict
            _writer.WriteEndElement();
        }
    }
}
