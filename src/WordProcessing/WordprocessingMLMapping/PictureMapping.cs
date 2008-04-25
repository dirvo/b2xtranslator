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

            _writer.WriteStartElement("w", "drawing", OpenXmlNamespaces.WordprocessingML);

            if (shape.fHaveAnchor)
            {
                _writer.WriteStartElement("wp", "anchor", OpenXmlNamespaces.WordprocessingDrawingML);
            }
            else
            {
                _writer.WriteStartElement("wp", "inline", OpenXmlNamespaces.WordprocessingDrawingML);
            }

            //convert the picture
            _writer.WriteStartElement("a", "graphic", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "graphicData", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("uri", OpenXmlNamespaces.DrawingMLPicture);
            _writer.WriteStartElement("pic", "pic", OpenXmlNamespaces.DrawingMLPicture);



            //write p:blipFil
            _writer.WriteStartElement("pic", "blipFill", OpenXmlNamespaces.DrawingMLPicture);

            _writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, _xmlPart.RelIdToString);
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingML);
            _writer.WriteElementString("a", "fillRect", OpenXmlNamespaces.DrawingML, "");
            _writer.WriteEndElement();

            //close p:blipFill
            _writer.WriteEndElement();



            //write p:spPr
            _writer.WriteStartElement("pic", "spPr", OpenXmlNamespaces.DrawingMLPicture);
            
            //write frame
            _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("x", "0");
            _writer.WriteAttributeString("y", "0");
            _writer.WriteEndElement();
            _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("cx", "3810000");
            _writer.WriteAttributeString("cy", "3810000");
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            //write preset geometry
            _writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("prst", "rect");
            _writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
            _writer.WriteEndElement();

            //close p:spPr
            _writer.WriteEndElement();




            //close p:pic
            _writer.WriteEndElement();

            //close a:graphic and a:graphicData
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            //close wp:anchor or wp:inline
            _writer.WriteEndElement();

            //close w:drawing
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
