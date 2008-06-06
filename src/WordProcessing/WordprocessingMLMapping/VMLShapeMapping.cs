using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.IO;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class VMLShapeMapping: PropertiesMapping,
          IMapping<FileShapeAddress>
    {
        ContentPart _targetPart;
        XmlElement _shapeNode, _imagedata;

        public VMLShapeMapping(XmlWriter writer, ContentPart targetPart)
            : base(writer)
        {
            _shapeNode = _nodeFactory.CreateElement("v", "shape", OpenXmlNamespaces.VectorML);
            _imagedata = _nodeFactory.CreateElement("v", "imagedata", OpenXmlNamespaces.VectorML);
        }

        public void Apply(FileShapeAddress fspa)
        {
            Shape shape = (Shape)fspa.ShapeContainer.Children[0];
            List<ShapeOptions.OptionEntry> options = getOptions(fspa.ShapeContainer);

            _writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

            //append id
            appendValueAttribute(_shapeNode, null, "id", getShapeId(shape), OpenXmlNamespaces.VectorML);

            //build style
            StringBuilder style = new StringBuilder();

            //append size and position ...
            TwipsValue left = new TwipsValue(fspa.xaLeft);
            TwipsValue top = new TwipsValue(fspa.yaTop);
            TwipsValue width = new TwipsValue(fspa.xaRight - fspa.xaLeft);
            TwipsValue height = new TwipsValue(fspa.yaBottom - fspa.yaTop);
            appendStyleProperty(style, "position", "absolute");
            appendStyleProperty(style, "margin-left", left.ToPoints()+"pt");
            appendStyleProperty(style, "margin-top", top.ToPoints() + "pt");
            appendStyleProperty(style, "width", width.ToPoints() + "pt");
            appendStyleProperty(style, "height", height.ToPoints() + "pt");

            foreach (ShapeOptions.OptionEntry entry in options)
            {
                switch (entry.pid)  
                {
                    case ShapeOptions.PropertyId.pWrapPolygonVertices:
                        appendValueAttribute(_shapeNode, null, "wrapcoords", getWrapCoords(entry), OpenXmlNamespaces.VectorML);
                        break;

                    case ShapeOptions.PropertyId.Pib:
                        string relId = copyImage(entry.op);
                        appendValueAttribute(_imagedata, "r", "id", relId, OpenXmlNamespaces.Relationships);
                        break;

                    case ShapeOptions.PropertyId.pibName:
                        string name = Encoding.Unicode.GetString(entry.opComplex);
                        name = name.Substring(0, name.Length - 1);
                        appendValueAttribute(_imagedata, "o", "title", name, OpenXmlNamespaces.Office);
                        break;
                }
            }

            //append the style
            appendValueAttribute(_shapeNode, null, "style", style.ToString(), OpenXmlNamespaces.VectorML);

            //append imagedata
            //if (_imagedata.Attributes.Count > 0)
            //    _shapeNode.AppendChild(_imagedata);

            //append wrap
            XmlElement wrap = _nodeFactory.CreateElement("w10", "wrap", OpenXmlNamespaces.OfficeWord);
            appendValueAttribute(wrap, null, "type", getWrapType(fspa), OpenXmlNamespaces.OfficeWord);
            _shapeNode.AppendChild(wrap);

            //write the shape
            _shapeNode.WriteTo(_writer);

            _writer.WriteEndElement();
            _writer.Flush();
        }


        private void appendStyleProperty(StringBuilder b, string propName, string propValue)
        {
            b.Append(propName);
            b.Append(":");
            b.Append(propValue);
            b.Append(";");
        }


        /// <summary>
        /// returns the OpenXML wrap type of the shape
        /// </summary>
        /// <param name="fspa"></param>
        /// <returns></returns>
        private string getWrapType(FileShapeAddress fspa)
        {
            switch (fspa.wr)
            {
                case 0:
                case 2:
                case 4:
                case 5:
                    return "tight";
                    break;
                case 1:
                    return "none";
                    break;
                case 3:
                    return "through";
                    break;
		        default:
                    return "none";
                    break;
	        }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="pib"></param>
        /// <returns></returns>
        private string copyImage(UInt32 pib)
        {
            return "";
        }


        /// <summary>
        /// Searches all OptionEntry in the ShapeContainer and puts them into a list.
        /// </summary>
        /// <param name="shapeContainer">The ShapeContainer</param>
        /// <returns>A List containing all OptionEntry of the ShapeContainer</returns>
        private List<ShapeOptions.OptionEntry> getOptions(ShapeContainer shapeContainer)
        {
            List<ShapeOptions.OptionEntry> ret = new List<ShapeOptions.OptionEntry>();

            //build the list of all option entries of this shape
            foreach (Record rec in shapeContainer.Children)
            {
                if (rec.GetType() == typeof(ShapeOptions))
                {
                    ShapeOptions opt = (ShapeOptions)rec;
                    ret.AddRange(opt.Options);
                }
            }

            return ret;
        }


        /// <summary>
        /// Generates a string id for the given shape
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private string getShapeId(Shape shape)
        {
            StringBuilder id = new StringBuilder();
            id.Append("_x0000_s");
            id.Append(shape.spid);
            return id.ToString();
        }


        /// <summary>
        /// Build the VML wrapcoords string for a given pWrapPolygonVertices
        /// </summary>
        /// <param name="pWrapPolygonVertices"></param>
        /// <returns></returns>
        private string getWrapCoords(ShapeOptions.OptionEntry pWrapPolygonVertices)
        {
            BinaryReader r = new BinaryReader(new MemoryStream(pWrapPolygonVertices.opComplex));
            List<Int32> pVertices = new List<Int32>();

            //skip first 6 bytes (header???)
            r.ReadBytes(6);

            //read the Int32 coordinates
            while (r.BaseStream.Position < r.BaseStream.Length)
            {
                pVertices.Add(r.ReadInt32());
            }

            //build the string
            StringBuilder coords = new StringBuilder();
            foreach (Int32 coord in pVertices)
            {
                coords.Append(coord);
                coords.Append(" ");
            }

            return coords.ToString().Trim();
        }
    }
}
