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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

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
            Shape shape = findShape(pict.ShapeContainer);

            _writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

            //v:shape
            _writer.WriteStartElement("v", "shape", OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("v", "type", OpenXmlNamespaces.VectorML, "rect");

            CultureInfo en = new CultureInfo("en-US");
            StringBuilder style = new StringBuilder();
            style.Append("width:").Append(new TwipsValue(pict.dxaGoal).ToPoints().ToString(en)).Append("pt;");
            style.Append("height:").Append(new TwipsValue(pict.dyaGoal).ToPoints().ToString(en)).Append("pt;");
            _writer.WriteAttributeString("style", style.ToString());

            //v:imageData
            _writer.WriteStartElement("v", "imageData", OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, _xmlPart.RelIdToString);
            _writer.WriteEndElement();

            //close v:shape
            _writer.WriteEndElement();

            //close w:pict
            _writer.WriteEndElement();
        }

        /// <summary>
        /// Finds the first Shape in the ShapeContainer 
        /// </summary>
        private Shape findShape(ShapeContainer con)
        {
            Shape ret = null;

            foreach (Record rec in con.Children)
            {
                if (rec.GetType() == typeof(Shape))
                {
                    ret = (Shape)rec;
                    break;
                }
            }

            return ret;
        }
    }
}
