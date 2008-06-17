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
          IMapping<ShapeContainer>
    {
        private BlipStoreContainer _blipStore = null;
        private ConversionContext _ctx;
        private FileShapeAddress _fspa;
        private ContentPart _targetPart;
        private XmlElement _fill, _stroke, _shadow, _imagedata;
        private bool _documentBase;


        public VMLShapeMapping(XmlWriter writer, ContentPart targetPart, FileShapeAddress fspa, bool documentBase, ConversionContext ctx)
            : base(writer)
        {
            _ctx = ctx;
            _fspa = fspa;
            _documentBase = documentBase;
            _targetPart = targetPart;
            _imagedata = _nodeFactory.CreateElement("v", "imagedata", OpenXmlNamespaces.VectorML);
            _fill = _nodeFactory.CreateElement("v", "fill", OpenXmlNamespaces.VectorML);
            _stroke = _nodeFactory.CreateElement("v", "stroke", OpenXmlNamespaces.VectorML);
            _shadow = _nodeFactory.CreateElement("v", "shadow", OpenXmlNamespaces.VectorML);

            Record recBs = _ctx.Doc.DrawingObjectTable.drawingGroup.FirstChildWithType<BlipStoreContainer>();
            if (recBs != null)
                _blipStore = (BlipStoreContainer)recBs;
        }


        public void Apply(ShapeContainer container)
        {
            if(_documentBase)
                _writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

            Record firstRecord = container.Children[0];
            if (firstRecord.GetType() == typeof(Shape))
            {
                //It's a single shape
                convertShape(container);
            }
            else if (firstRecord.GetType() == typeof(GroupShapeRecord))
            { 
                //Its a group of shapes
                convertGroup((GroupContainer)container.ParentRecord);
            }

            if(_documentBase)
                _writer.WriteEndElement();

            _writer.Flush();
        }


        /// <summary>
        /// Converts a group of shapes
        /// </summary>
        /// <param name="container"></param>
        private void convertGroup(GroupContainer container)
        {
            ShapeContainer groupShape = (ShapeContainer)container.Children[0];
            GroupShapeRecord gsr = (GroupShapeRecord)groupShape.Children[0];
            Shape shape = (Shape)groupShape.Children[1];
            List<ShapeOptions.OptionEntry> options = groupShape.ExtractOptions();

            _writer.WriteStartElement("v", "group", OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("id", getShapeId(shape));
            _writer.WriteAttributeString("coordorigin", gsr.rcgBounds.Left + "," + gsr.rcgBounds.Top);
            _writer.WriteAttributeString("coordsize", gsr.rcgBounds.Width + "," + gsr.rcgBounds.Height);

            if (_documentBase)
            {
                //this group is placed directly in the document, 
                //so use the FSPA to build the style
                StringBuilder style = buildStyle(_fspa);

                //write wrap coords
                foreach (ShapeOptions.OptionEntry entry in options)
                {
                    switch (entry.pid)
                    {
                        case ShapeOptions.PropertyId.pWrapPolygonVertices:
                            _writer.WriteAttributeString("wrapcoords", getWrapCoords(entry));
                            break;
                        case ShapeOptions.PropertyId.dhgt:
                            appendStyleProperty(style, "z-index", entry.op.ToString());
                            break;
                    }
                }

                _writer.WriteAttributeString("style", style.ToString());
            }

            //convert the shapes/groups in the group
            for (int i = 1; i < container.Children.Count; i++)
            {
                if (container.Children[i].GetType() == typeof(ShapeContainer))
                {
                    ShapeContainer childShape = (ShapeContainer)container.Children[i];
                    childShape.Convert(new VMLShapeMapping(_writer, _targetPart, _fspa, false, _ctx));
                }
                else if (container.Children[i].GetType() == typeof(GroupContainer))
                {
                    GroupContainer childGroup = (GroupContainer)container.Children[i];
                    convertGroup(childGroup);
                }
            }

            //write wrap
            if (_documentBase)
            {
                _writer.WriteStartElement("w10", "wrap", OpenXmlNamespaces.OfficeWord);
                _writer.WriteAttributeString("type", getWrapType(_fspa));
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
        }


        /// <summary>
        /// Converts a single shape
        /// </summary>
        /// <param name="container"></param>
        private void convertShape(ShapeContainer container)
        {
            Shape shape = (Shape)container.Children[0];
            List<ShapeOptions.OptionEntry> options = container.ExtractOptions();

            //write the shapeType
            if (shape.ShapeType != null)
            {
                shape.ShapeType.Convert(new VMLShapeTypeMapping(_writer));
            }

            _writer.WriteStartElement("v", "shape", OpenXmlNamespaces.VectorML);

            //append id
            _writer.WriteAttributeString("id", getShapeId(shape));

            //build the style
            StringBuilder style = null;
            if (_documentBase)
            {
                //this shape is placed directly in the document, 
                //so use the FSPA to build the style
                style = buildStyle(_fspa);
            }
            else
            {
                //use the anchor to build the style
                Record recAnchor = container.FirstChildWithType<ChildAnchor>();
                if (recAnchor != null)
                {
                    ChildAnchor anchor = (ChildAnchor)recAnchor;
                    style = buildStyle(anchor);
                }
            }

            EmuValue shadowOffsetX = null;
            EmuValue shadowOffsetY = null;
            double shadowOriginX = 0;
            double shadowOriginY = 0;

            foreach (ShapeOptions.OptionEntry entry in options)
            {
                switch (entry.pid)
                {
                    case ShapeOptions.PropertyId.pWrapPolygonVertices:
                        _writer.WriteAttributeString("wrapcoords", getWrapCoords(entry));
                        break;

                    // GEOMETRY

                    case ShapeOptions.PropertyId.rotation:
                        appendStyleProperty(style, "rotation", (entry.op / Math.Pow(2, 16)).ToString());
                        break;

                    // OUTLINE

                    case ShapeOptions.PropertyId.lineColor:
                        RGBColor lineColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        _writer.WriteAttributeString("strokecolor", "#" + lineColor.SixDigitHexCode);
                        break;

                    case ShapeOptions.PropertyId.lineWidth:
                        EmuValue lineWidth = new EmuValue((int)entry.op);
                        _writer.WriteAttributeString("strokeweight", lineWidth.ToPoints() + "pt");
                        break;

                    case ShapeOptions.PropertyId.lineDashing:
                        Global.DashStyle dash = (Global.DashStyle)entry.op;
                        appendValueAttribute(_stroke, null, "dashstyle", dash.ToString(), null);
                        break;

                    case ShapeOptions.PropertyId.lineStyle:
                        appendValueAttribute(_stroke, null, "linestyle", getLineStyle(entry.op), null);
                        break;

                    // FILL

                    case ShapeOptions.PropertyId.fillColor:
                        RGBColor fillColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        _writer.WriteAttributeString("fillcolor", "#" + fillColor.SixDigitHexCode);
                        break;

                    case ShapeOptions.PropertyId.fillBackColor:
                        RGBColor fillBackColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        appendValueAttribute(_fill, null, "color2", "#" + fillBackColor.SixDigitHexCode, null);
                        break;

                    case ShapeOptions.PropertyId.fillShadeType:
                        appendValueAttribute(_fill, null, "method", getFillMethod(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.fillFocus:
                        appendValueAttribute(_fill, null, "focus", entry.op + "%", null);
                        break;

                    case ShapeOptions.PropertyId.fillType:
                        appendValueAttribute(_fill, null, "type", getFillType(entry.op), null);
                        break;

                    // SHADOW

                    case ShapeOptions.PropertyId.shadowType:
                        appendValueAttribute(_shadow, null, "type", getShadowType(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.shadowColor:
                        RGBColor shadowColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        appendValueAttribute(_shadow, null, "color", "#" + shadowColor.SixDigitHexCode, null);
                        break;

                    case ShapeOptions.PropertyId.shadowOffsetX:
                        shadowOffsetX = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowOffsetY:
                        shadowOffsetY = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowOriginX:
                        shadowOriginX = entry.op / Math.Pow(2, 16);
                        break;

                    case ShapeOptions.PropertyId.shadowOriginY:
                        shadowOriginY = entry.op / Math.Pow(2, 16);
                        break;

                    case ShapeOptions.PropertyId.shadowOpacity:
                        appendValueAttribute(_shadow, null, "opacity", (entry.op / Math.Pow(2, 16)).ToString(), null);
                        break;

                    // PICTURE
                    
                    case ShapeOptions.PropertyId.Pib:
                        int index = (int)entry.op - 1;
                        BlipStoreEntry bse = (BlipStoreEntry)_blipStore.Children[index];
                        ImagePart part = copyPicture(bse);
                        appendValueAttribute(_imagedata, "r", "id", part.RelIdToString, OpenXmlNamespaces.Relationships);
                        break;

                    case ShapeOptions.PropertyId.pibName:
                        string name = Encoding.Unicode.GetString(entry.opComplex);
                        name = name.Substring(0, name.Length - 1);
                        appendValueAttribute(_imagedata, "o", "title", name, OpenXmlNamespaces.Office);
                        break;
                }
            }

            //write type
            if (shape.ShapeType != null)
            {
                _writer.WriteAttributeString("type", "#" + VMLShapeTypeMapping.GenerateTypeId(shape.ShapeType));
            }

            //write the style
            if (style != null)
            {
                _writer.WriteAttributeString("style", style.ToString());
            }

            //build shadow offset
            if (shadowOffsetX != null && shadowOffsetY != null)
            {
                appendValueAttribute(
                    _shadow, null, "offset", 
                    shadowOffsetX.ToPoints() + "pt," + shadowOffsetY.ToPoints() + "pt", 
                    null);
            }

            //build shadow origin
            if (shadowOriginX != 0 && shadowOriginY != 0)
            {
                appendValueAttribute(
                    _shadow, null, "origin",
                    shadowOriginX + "," + shadowOriginY,
                    null);
            }

            //write shadow
            if (_shadow.Attributes.Count > 0)
            {
                appendValueAttribute(_shadow, null, "on", "t", null);
                _shadow.WriteTo(_writer);
            }

            //write wrap
            if (_documentBase)
            {
                _writer.WriteStartElement("w10", "wrap", OpenXmlNamespaces.OfficeWord);
                _writer.WriteAttributeString("type", getWrapType(_fspa));
                _writer.WriteEndElement();
            }

            //write stroke
            if (_stroke.Attributes.Count > 0)
            {
                _stroke.WriteTo(_writer);
            }

            //write fill
            if (_fill.Attributes.Count > 0)
            {
                _fill.WriteTo(_writer);
            }

            //write imagedata
            if (_imagedata.Attributes.Count > 0)
            {
                _imagedata.WriteTo(_writer);
            }

            //write the textbox
            Record recTextbox = container.FirstChildWithType<ClientTextbox>();
            if (recTextbox != null)
            {
                ClientTextbox txbx = (ClientTextbox)recTextbox;
                _ctx.Doc.Convert(new TextboxMapping(_ctx, txbx, _targetPart, _writer));
            }

            //write the shape
            _writer.WriteEndElement();
            _writer.Flush();
        }


        /// <summary>
        /// Returns the OpenXML fill type of a fill effect
        /// </summary>
        private string getFillType(uint p)
        {
            switch (p)
            {
                case 0:
                    return "solid";
                case 1:
                    return "tile";
                case 2:
                    return "pattern";
                case 3:
                    return "frame";
                case 4:
                    return "gradient";
                case 5:
                    return "gradientRadial";
                case 6:
                    return "gradientRadial";
                case 7:
                    return "gradient";
                case 9:
                    return "solid";
                default:
                    return "solid";
            }
        }


        private string getShadowType(uint p)
        {
            switch (p)
            {
                case 0:
                    return "single";
                case 1:
                    return "double";
                case 2:
                    return "perspective";
                case 3:
                    return "shaperelative";
                case 4:
                    return "drawingrelative";
                case 5:
                    return "emboss";
                default:
                    return "single";
            }
        }

        private string getLineStyle(uint p)
        {
            switch (p)
            {
                case 0:
                    return "single";
                case 1:
                    return "thinThin";
                case 2:
                    return "thinThick";
                case 3:
                    return "thickThin";
                case 4:
                    return "thickBetweenThin";
                default:
                    return "single";
            }
        }

        private string getFillMethod(uint p)
        {
            Int16 val = (Int16)((p & 0xFFFF0000) >> 28);
            switch (val)
            {
                case 0:
                    return "none";
                case 1:
                    return "any";
                case 2:
                    return "linear";
                case 4:
                    return "linear sigma";
                default:
                    return "any";
            }
        }


        /// <summary>
        /// Returns the OpenXML wrap type of the shape
        /// </summary>
        /// <param name="fspa"></param>
        /// <returns></returns>
        private string getWrapType(FileShapeAddress fspa)
        {
            switch (fspa.wr)
            {
                case 0:
                case 2:
                    return "square";
                case 1:
                    return "none";
                case 3:
                    return "through";
                case 4:
                case 5:
                    return "tight";

		        default:
                    return "none";
	        }
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


        private StringBuilder buildStyle(FileShapeAddress fspa)
        {
            //build style
            StringBuilder style = new StringBuilder();

            //append size and position ...
            TwipsValue left = new TwipsValue(fspa.xaLeft);
            TwipsValue top = new TwipsValue(fspa.yaTop);
            TwipsValue width = new TwipsValue(fspa.xaRight - fspa.xaLeft);
            TwipsValue height = new TwipsValue(fspa.yaBottom - fspa.yaTop);
            appendStyleProperty(style, "position", "absolute");
            appendStyleProperty(style, "margin-left", left.ToPoints() + "pt");
            appendStyleProperty(style, "margin-top", top.ToPoints() + "pt");
            appendStyleProperty(style, "width", width.ToPoints() + "pt");
            appendStyleProperty(style, "height", height.ToPoints() + "pt");
            appendStyleProperty(style, "mso-position-horizontal-relative", fspa.bx.ToString());
            appendStyleProperty(style, "mso-position-vertical-relative", fspa.by.ToString());

            return style;
        }

        private StringBuilder buildStyle(ChildAnchor anchor)
        {
            //build style
            StringBuilder style = new StringBuilder();

            //append size and position ...
            appendStyleProperty(style, "position", "absolute");
            appendStyleProperty(style, "left", anchor.rcgBounds.Left.ToString());
            appendStyleProperty(style, "top", anchor.rcgBounds.Top.ToString());
            appendStyleProperty(style, "width", anchor.rcgBounds.Width.ToString());
            appendStyleProperty(style, "height", anchor.rcgBounds.Height.ToString());

            return style;
        }


        private void appendStyleProperty(StringBuilder b, string propName, string propValue)
        {
            b.Append(propName);
            b.Append(":");
            b.Append(propValue);
            b.Append(";");
        }


        /// <summary>
        /// Copies the picture from the binary stream to the zip archive 
        /// and creates the relationships for the image.
        /// </summary>
        /// <param name="pict">The PictureDescriptor</param>
        /// <returns>The created ImagePart</returns>
        protected ImagePart copyPicture(BlipStoreEntry bse)
        {
            //create the image part
            ImagePart imgPart = null;

            switch (bse.btWin32)
            {
                case BlipStoreEntry.BlipType.msoblipEMF:
                    imgPart = _targetPart.AddImagePart(ImagePartType.Emf);
                    break;
                case BlipStoreEntry.BlipType.msoblipWMF:
                    imgPart = _targetPart.AddImagePart(ImagePartType.Wmf);
                    break;
                case BlipStoreEntry.BlipType.msoblipJPEG:
                case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                    imgPart = _targetPart.AddImagePart(ImagePartType.Jpeg);
                    break;
                case BlipStoreEntry.BlipType.msoblipPNG:
                    imgPart = _targetPart.AddImagePart(ImagePartType.Png);
                    break;
                case BlipStoreEntry.BlipType.msoblipTIFF:
                    imgPart = _targetPart.AddImagePart(ImagePartType.Tiff);
                    break;
                case BlipStoreEntry.BlipType.msoblipPICT:
                case BlipStoreEntry.BlipType.msoblipDIB:
                case BlipStoreEntry.BlipType.msoblipERROR:
                case BlipStoreEntry.BlipType.msoblipUNKNOWN:
                case BlipStoreEntry.BlipType.msoblipLastClient:
                case BlipStoreEntry.BlipType.msoblipFirstClient:
                    throw new MappingException("Cannot convert picture of type " + bse.btWin32);
            }

            Stream outStream = imgPart.GetStream();

            _ctx.Doc.WordDocumentStream.Seek((long)bse.foDelay, SeekOrigin.Begin);
            BinaryReader reader = new BinaryReader(_ctx.Doc.WordDocumentStream);

            switch (bse.btWin32)
            {
                case BlipStoreEntry.BlipType.msoblipEMF:
                case BlipStoreEntry.BlipType.msoblipWMF:

                    //it's a meta image
                    MetafilePictBlip metaBlip = (MetafilePictBlip)Record.readRecord(reader);

                    //meta images can be compressed
                    byte[] decompressed = metaBlip.Decrompress();
                    outStream.Write(decompressed, 0, decompressed.Length);

                    break;
                case BlipStoreEntry.BlipType.msoblipJPEG:
                case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                case BlipStoreEntry.BlipType.msoblipPNG:
                case BlipStoreEntry.BlipType.msoblipTIFF:

                    //it's a bitmap image
                    BitmapBlip bitBlip = (BitmapBlip)Record.readRecord(reader);
                    outStream.Write(bitBlip.m_pvBits, 0, bitBlip.m_pvBits.Length);
                    break;
            }

            return imgPart;
        }
    }
}
