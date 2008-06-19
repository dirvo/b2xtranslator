using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class VMLShapeTypeMapping : PropertiesMapping,
          IMapping<ShapeType>
    {
        
        public VMLShapeTypeMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(ShapeType shapeType)
        {
            _writer.WriteStartElement("v", "shapetype", OpenXmlNamespaces.VectorML);

            //id
            _writer.WriteAttributeString("id", GenerateTypeId(shapeType));

            //coordinate system
            _writer.WriteAttributeString("coordsize", "21600,21600");

            //shape type code
            _writer.WriteAttributeString("o", "spt", OpenXmlNamespaces.Office, shapeType.TypeCode.ToString());

            //adj
            _writer.WriteAttributeString("adj", "10800");

            //The path
            if (shapeType.Path != null)
                _writer.WriteAttributeString("path", shapeType.Path);

            //stroke
            _writer.WriteStartElement("v", "stroke", OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("joinstyle", shapeType.Joins.ToString());
            _writer.WriteEndElement();

            //Formulas
            if (shapeType.Formulas != null && shapeType.Formulas.Count > 0)
            {
                _writer.WriteStartElement("v", "formulas", OpenXmlNamespaces.VectorML);

                foreach (string formula in shapeType.Formulas)
                {
                    _writer.WriteStartElement("v", "f", OpenXmlNamespaces.VectorML);
                    _writer.WriteAttributeString("eqn", formula);
                    _writer.WriteEndElement();
                }

                _writer.WriteEndElement();
            }


            //Path
            _writer.WriteStartElement("v", "path", OpenXmlNamespaces.VectorML);
            if(shapeType.ShapeConcentricFill)
            {
                _writer.WriteAttributeString("gradientshapeok", "t");
            }
            if (shapeType.Limo != null)
            {
                _writer.WriteAttributeString("limo", shapeType.Limo);
            }
            if (shapeType.ConnectorLocations != null)
            {
                _writer.WriteAttributeString("o", "connecttype", OpenXmlNamespaces.Office, "custom");
                _writer.WriteAttributeString("o", "connectlocs", OpenXmlNamespaces.Office, shapeType.ConnectorLocations);
            }
            if (shapeType.TextboxRectangle != null)
            {
                _writer.WriteAttributeString("textboxrect", shapeType.TextboxRectangle);
            }
            _writer.WriteEndElement();


            //Handles
            if (shapeType.Handles != null && shapeType.Handles.Count > 0)
            {
                _writer.WriteStartElement("v", "handles", OpenXmlNamespaces.VectorML);

                foreach (ShapeType.Handle handle in shapeType.Handles)
                {
                    _writer.WriteStartElement("v", "h", OpenXmlNamespaces.VectorML);

                    if(handle.position != null)
                        _writer.WriteAttributeString("position", handle.position);
                    
                    if (handle.switchHandle != null)
                        _writer.WriteAttributeString("switch", handle.switchHandle);
                    
                    if (handle.xrange != null)
                        _writer.WriteAttributeString("xrange", handle.xrange);

                    if (handle.yrange != null)
                        _writer.WriteAttributeString("yrange", handle.yrange);

                    
                    _writer.WriteEndElement();
                }

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
        }


        /// <summary>
        /// Returns the id of the referenced type
        /// </summary>
        public static string GenerateTypeId(ShapeType shapeType)
        {
            StringBuilder type = new StringBuilder();
            type.Append("_x0000_t");
            type.Append(shapeType.TypeCode);
            return type.ToString();
        }

        
    }
}
