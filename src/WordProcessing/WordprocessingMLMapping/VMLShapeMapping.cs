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
using System.Globalization;
using DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class VMLShapeMapping: PropertiesMapping,
          IMapping<ShapeContainer>
    {
        private BlipStoreContainer _blipStore = null;
        private ConversionContext _ctx;
        private FileShapeAddress _fspa;
        private ContentPart _targetPart;
        private XmlElement _fill, _stroke, _shadow, _imagedata, _3dstyle;
        private bool _documentBase;
        private List<byte> pSegmentInfo = new List<byte>();
        private List<byte> pVertices = new List<byte>();

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
            _3dstyle = _nodeFactory.CreateElement("o", "extrusion", OpenXmlNamespaces.Office);

            Record recBs = _ctx.Doc.OfficeArtContent.DrawingGroupData.FirstChildWithType<BlipStoreContainer>();
            if (recBs != null)
                _blipStore = (BlipStoreContainer)recBs;
        }


        public void Apply(ShapeContainer container)
        {
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
            ChildAnchor anchor = groupShape.FirstChildWithType<ChildAnchor>();

            _writer.WriteStartElement("v", "group", OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("id", getShapeId(shape));
            _writer.WriteAttributeString("style", generateStyle(anchor, options).ToString());
            _writer.WriteAttributeString("coordorigin", gsr.rcgBounds.Left + "," + gsr.rcgBounds.Top);
            _writer.WriteAttributeString("coordsize", gsr.rcgBounds.Width + "," + gsr.rcgBounds.Height);
            
            //write wrap coords
            foreach (ShapeOptions.OptionEntry entry in options)
            {
                switch (entry.pid)
                {
                    case ShapeOptions.PropertyId.pWrapPolygonVertices:
                        _writer.WriteAttributeString("wrapcoords", getWrapCoords(entry));
                        break;
                }
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
                    _documentBase = false;
                    convertGroup(childGroup);
                }
            }

            //write wrap
            if (_documentBase)
            {
                string wrap = getWrapType(_fspa);
                if(wrap != "through")
                {
                    _writer.WriteStartElement("w10", "wrap", OpenXmlNamespaces.OfficeWord);
                    _writer.WriteAttributeString("type", wrap);
                    _writer.WriteEndElement();
                }
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
            ChildAnchor anchor = container.FirstChildWithType<ChildAnchor>();
            ClientAnchor clientAnchor = container.FirstChildWithType<ClientAnchor>();

            writeStartShapeElement(shape);
            _writer.WriteAttributeString("id", getShapeId(shape));
            if (shape.ShapeType != null)
            {
                _writer.WriteAttributeString("type", "#" + VMLShapeTypeMapping.GenerateTypeId(shape.ShapeType));
            }
            _writer.WriteAttributeString("style", generateStyle(anchor, options).ToString());

            //temporary variables
            EmuValue shadowOffsetX = null;
            EmuValue shadowOffsetY = null;
            EmuValue secondShadowOffsetX = null;
            EmuValue secondShadowOffsetY = null;
            double shadowOriginX = 0;
            double shadowOriginY = 0;
            string[] adjValues = new string[8];
            int numberAdjValues = 0;
            uint xCoord = 0;
            uint yCoord = 0;
            bool stroked = true;
            bool filled = true;
            bool hasTextbox = false;

            foreach (ShapeOptions.OptionEntry entry in options)
            {
                switch (entry.pid)
                {
                    //BOOLEANS

                    case ShapeOptions.PropertyId.geometryBooleans:
                        GeometryBooleans geometryBooleans = new GeometryBooleans(entry.op);

                        if (!(geometryBooleans.fUsefLineOK && geometryBooleans.fLineOK))
                        {
                            stroked = false;
                        }

                        if (!(geometryBooleans.fUsefFillOK && geometryBooleans.fFillOK))
                        {
                            filled = false;
                        }
                        break;

                    case ShapeOptions.PropertyId.lineStyleBooleans:
                        LineStyleBooleans lineBooleans = new LineStyleBooleans(entry.op);

                        if (!(lineBooleans.fUsefLine && lineBooleans.fLine))
                        {
                            stroked = false;
                        }

                        break;

                    case ShapeOptions.PropertyId.protectionBooleans:
                        ProtectionBooleans protBools = new ProtectionBooleans(entry.op);

                        break;

                    case ShapeOptions.PropertyId.diagramBooleans:
                        DiagramBooleans diaBools = new DiagramBooleans(entry.op);

                        break;

                    // GEOMETRY

                    case ShapeOptions.PropertyId.adjustValue:
                        adjValues[0] = (((int)entry.op).ToString());
                        numberAdjValues++; 
                        break;

                    case ShapeOptions.PropertyId.adjust2Value:
                        adjValues[1] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust3Value:
                        adjValues[2] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust4Value:
                        adjValues[3] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust5Value:
                        adjValues[4] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust6Value:
                        adjValues[5] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust7Value:
                        adjValues[6] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust8Value:
                        adjValues[7] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.pWrapPolygonVertices:
                        _writer.WriteAttributeString("wrapcoords", getWrapCoords(entry));
                        break;

                    case ShapeOptions.PropertyId.geoRight:
                        xCoord = entry.op;
                        break;

                    case ShapeOptions.PropertyId.geoBottom:
                        yCoord = entry.op;
                        break;

                    // OUTLINE

                    case ShapeOptions.PropertyId.lineColor:
                        RGBColor lineColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        _writer.WriteAttributeString("strokecolor", "#" + lineColor.SixDigitHexCode);
                        break;

                    case ShapeOptions.PropertyId.lineWidth:
                        EmuValue lineWidth = new EmuValue((int)entry.op);
                        _writer.WriteAttributeString("strokeweight", lineWidth.ToString());
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

                    case ShapeOptions.PropertyId.fillBlip:
                        BlipStoreEntry fillBlip = (BlipStoreEntry)_blipStore.Children[(int)entry.op - 1];
                        ImagePart fillBlipPart = copyPicture(fillBlip);
                        appendValueAttribute(_fill, "r", "id", fillBlipPart.RelIdToString, OpenXmlNamespaces.Relationships);
                        appendValueAttribute(_imagedata, "o", "title", "", OpenXmlNamespaces.Office);
                        break;

                    case ShapeOptions.PropertyId.fillOpacity:
                        appendValueAttribute(_fill, null, "opacity", entry.op + "f" , null);
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

                    case ShapeOptions.PropertyId.shadowSecondOffsetX:
                        secondShadowOffsetX = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowOffsetY:
                        shadowOffsetY = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowSecondOffsetY:
                        secondShadowOffsetY = new EmuValue((int)entry.op);
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

                    // 3D STYLE

                    case ShapeOptions.PropertyId.f3D:
                    case ShapeOptions.PropertyId.fc3DFillHarsh:
                    case ShapeOptions.PropertyId.fc3DLightFace:
                        break;
                    case ShapeOptions.PropertyId.c3DExtrudeBackward:
                        EmuValue backwardValue = new EmuValue((int)entry.op);
                        appendValueAttribute(_3dstyle, "backdepth", backwardValue.ToPoints().ToString());
                        break; 

                    // TEXTBOX

                    case ShapeOptions.PropertyId.lTxid:
                        hasTextbox = true;
                        break;

                    // PATH
                    case ShapeOptions.PropertyId.shapePath:
                        //
                        _writer.WriteAttributeString("path", parsePath(options));
                        break;
                }
            }

            if (!filled)
            {
                _writer.WriteAttributeString("filled", "f");
            }

            if (!stroked)
            {
                _writer.WriteAttributeString("stroked", "f");
            }

            if (xCoord > 0 && yCoord > 0)
            {
                _writer.WriteAttributeString("coordsize", xCoord + "," + yCoord);
            }

            //write adj values 
            if (numberAdjValues != 0)
            {
                string adjString = adjValues[0];
                for (int i = 1; i < 8; i++)
                {
                    adjString += "," + adjValues[i];
                }
                _writer.WriteAttributeString("adj", adjString);
                //string.Format("{0:x4}", adjValues);
            }

            //build shadow offsets
            StringBuilder offset = new StringBuilder();
            if (shadowOffsetX != null)
            {
                offset.Append(shadowOffsetX.ToPoints());
                offset.Append("pt");
            }
            if (shadowOffsetY != null)
            {
                offset.Append(",");
                offset.Append(shadowOffsetY.ToPoints());
                offset.Append("pt");
            }
            if (offset.Length > 0)
            {
                appendValueAttribute(_shadow, null, "offset", offset.ToString(), null);
            }
            StringBuilder offset2 = new StringBuilder();
            if (secondShadowOffsetX != null)
            {
                offset2.Append(secondShadowOffsetX.ToPoints());
                offset2.Append("pt");
            }
            if (secondShadowOffsetY != null)
            {
                offset2.Append(",");
                offset2.Append(secondShadowOffsetY.ToPoints());
                offset2.Append("pt");
            }
            if (offset2.Length > 0)
            {
                appendValueAttribute(_shadow, null, "offset2", offset2.ToString(), null);
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

            //write 3d style 
            if (_3dstyle.Attributes.Count > 0)
            {
                appendValueAttribute(_3dstyle, "v", "ext", "view", OpenXmlNamespaces.VectorML);
                appendValueAttribute(_3dstyle, null, "on", "t", null);
                _3dstyle.WriteTo(_writer);
            }

            //write wrap
            if (_documentBase)
            {
                string wrap = getWrapType(_fspa);
                if(wrap != "through")
                {
                    _writer.WriteStartElement("w10", "wrap", OpenXmlNamespaces.OfficeWord);
                    _writer.WriteAttributeString("type", wrap);
                    _writer.WriteEndElement();
                }
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
                //Word text box

                //Word appends a ClientTextbox record to the container. 
                //This record stores the index of the textbox.

                ClientTextbox box = (ClientTextbox)recTextbox;
                Int16 textboxIndex = System.BitConverter.ToInt16(box.Bytes, 2);
                textboxIndex--;
                _ctx.Doc.Convert(new TextboxMapping(_ctx, textboxIndex, _targetPart, _writer));
            }
            else if(hasTextbox)
            {
                //Open Office textbox

                //Open Office doesn't append a ClientTextbox record to the container.
                //We don't know how Word gets the relation to the text, but we assume that the first textbox in the document
                //get the index 0, the second textbox gets the index 1 (and so on).

                _ctx.Doc.Convert(new TextboxMapping(_ctx, _targetPart, _writer));
            }

            //write the shape
            _writer.WriteEndElement();
            _writer.Flush();
        }

        private string parsePath(List<ShapeOptions.OptionEntry> options)
        {
            string path = "";
            byte[] pVertices = null;
            byte[] pSegmentInfo = null;

            foreach (ShapeOptions.OptionEntry e in options)
            {
                if (e.pid == ShapeOptions.PropertyId.pVertices)
                {
                    pVertices = e.opComplex;
                }
                else if (e.pid == ShapeOptions.PropertyId.pSegmentInfo)
                {
                    pSegmentInfo = e.opComplex;
                }
            }

            if (pSegmentInfo != null && pVertices.Length != null)
            {
                PathParser parser = new PathParser(pSegmentInfo, pVertices);
                path = parser.VmlPath.ToString();
            }
            return path;
        }

        private StringBuilder generateStyle(ChildAnchor anchor, List<ShapeOptions.OptionEntry> options)
        {
            StringBuilder style = new StringBuilder();
            if (_documentBase)
            {
                //this shape is placed directly in the document, 
                //so use the FSPA to build the style
                AppendDimensionToStyle(style, _fspa);
            }
            else
            {
                //the style is part of a group, 
                //so use the anchor
                AppendDimensionToStyle(style, anchor);
            }
            AppendOptionsToStyle(style, options);
            return style;
        }

        private void writeStartShapeElement(Shape shape)
        {
            if (shape.ShapeType is OvalType)
            {
                //OVAL
                _writer.WriteStartElement("v", "oval", OpenXmlNamespaces.VectorML);
            }
            else if (shape.ShapeType is RoundedRectangleType)
            {
                //ROUNDED RECT
                _writer.WriteStartElement("v", "roundrect", OpenXmlNamespaces.VectorML);
            }
            else if (shape.ShapeType is RectangleType)
            {
                //RECT
                _writer.WriteStartElement("v", "rect", OpenXmlNamespaces.VectorML);
            }
            else
            {
                //SHAPE
                if (shape.ShapeType != null)
                {
                    shape.ShapeType.Convert(new VMLShapeTypeMapping(_writer));
                }
                _writer.WriteStartElement("v", "shape", OpenXmlNamespaces.VectorML);
            }
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
                    imgPart = _targetPart.AddImagePart(ImagePart.ImageType.Emf);
                    break;
                case BlipStoreEntry.BlipType.msoblipWMF:
                    imgPart = _targetPart.AddImagePart(ImagePart.ImageType.Wmf);
                    break;
                case BlipStoreEntry.BlipType.msoblipJPEG:
                case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                    imgPart = _targetPart.AddImagePart(ImagePart.ImageType.Jpeg);
                    break;
                case BlipStoreEntry.BlipType.msoblipPNG:
                    imgPart = _targetPart.AddImagePart(ImagePart.ImageType.Png);
                    break;
                case BlipStoreEntry.BlipType.msoblipTIFF:
                    imgPart = _targetPart.AddImagePart(ImagePart.ImageType.Tiff);
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
                    MetafilePictBlip metaBlip = (MetafilePictBlip)Record.ReadRecord(reader, 0);

                    //meta images can be compressed
                    byte[] decompressed = metaBlip.Decrompress();
                    outStream.Write(decompressed, 0, decompressed.Length);

                    break;
                case BlipStoreEntry.BlipType.msoblipJPEG:
                case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                case BlipStoreEntry.BlipType.msoblipPNG:
                case BlipStoreEntry.BlipType.msoblipTIFF:

                    //it's a bitmap image
                    BitmapBlip bitBlip = (BitmapBlip)Record.ReadRecord(reader, 0);
                    outStream.Write(bitBlip.m_pvBits, 0, bitBlip.m_pvBits.Length);
                    break;
            }

            return imgPart;
        }

        //*******************************************************************
        //                                                     STATIC METHODS
        //*******************************************************************

        public static void AppendDimensionToStyle(StringBuilder style, FileShapeAddress fspa)
        {
            //append size and position ...
            TwipsValue left = new TwipsValue(fspa.xaLeft);
            TwipsValue top = new TwipsValue(fspa.yaTop);
            TwipsValue width = new TwipsValue(fspa.xaRight - fspa.xaLeft);
            TwipsValue height = new TwipsValue(fspa.yaBottom - fspa.yaTop);

            appendStyleProperty(style, "position", "absolute");
            appendStyleProperty(style, "margin-left", Convert.ToString(left.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
            appendStyleProperty(style, "margin-top", Convert.ToString(top.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
            appendStyleProperty(style, "width", Convert.ToString(width.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
            appendStyleProperty(style, "height", Convert.ToString(height.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
        }

        public static void AppendDimensionToStyle(StringBuilder style, ChildAnchor anchor)
        {
            //append size and position ...
            appendStyleProperty(style, "position", "absolute");
            appendStyleProperty(style, "left", anchor.rcgBounds.Left.ToString());
            appendStyleProperty(style, "top", anchor.rcgBounds.Top.ToString());
            appendStyleProperty(style, "width", anchor.rcgBounds.Width.ToString());
            appendStyleProperty(style, "height", anchor.rcgBounds.Height.ToString());
        }

        public static void AppendOptionsToStyle(StringBuilder style, List<ShapeOptions.OptionEntry> options)
        {
            foreach (ShapeOptions.OptionEntry entry in options)
            {
                switch (entry.pid)
                {

                    //POSITIONING

                    case ShapeOptions.PropertyId.posh:
                        appendStyleProperty(
                            style,
                            "mso-position-horizontal",
                            mapHorizontalPosition((ShapeOptions.PositionHorizontal)entry.op));
                        break;
                    case ShapeOptions.PropertyId.posrelh:
                        appendStyleProperty(
                            style,
                            "mso-position-horizontal-relative",
                            mapHorizontalPositionRelative((ShapeOptions.PositionHorizontalRelative)entry.op));
                        break;
                    case ShapeOptions.PropertyId.posv:
                        appendStyleProperty(
                            style,
                            "mso-position-vertical",
                            mapVerticalPosition((ShapeOptions.PositionVertical)entry.op));
                        break;
                    case ShapeOptions.PropertyId.posrelv:
                        appendStyleProperty(
                            style,
                            "mso-position-vertical-relative",
                            mapVerticalPositionRelative((ShapeOptions.PositionVerticalRelative)entry.op));
                        break;

                    //BOOLEANS

                    case ShapeOptions.PropertyId.groupShapeBooleans:
                        GroupShapeBooleans groupShapeBoolean = new GroupShapeBooleans(entry.op);

                        if (groupShapeBoolean.fUsefBehindDocument && groupShapeBoolean.fBehindDocument)
                        {
                            //The shape is behind the text, so the z-index must be negative.
                            appendStyleProperty(style, "z-index", "-1");
                        }

                        break;

                    // GEOMETRY

                    case ShapeOptions.PropertyId.rotation:
                        appendStyleProperty(style, "rotation", (entry.op / Math.Pow(2, 16)).ToString());
                        break;

                    //TEXTBOX

                    case ShapeOptions.PropertyId.anchorText:
                        appendStyleProperty(style, "v-text-anchor", getTextboxAnchor(entry.op));
                        break;

                    //WRAP DISTANCE

                    case ShapeOptions.PropertyId.dxWrapDistLeft:
                        appendStyleProperty(style, "mso-wrap-distance-left", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                    case ShapeOptions.PropertyId.dxWrapDistRight:
                        appendStyleProperty(style, "mso-wrap-distance-right", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                    case ShapeOptions.PropertyId.dyWrapDistBottom:
                        appendStyleProperty(style, "mso-wrap-distance-bottom", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                    case ShapeOptions.PropertyId.dyWrapDistTop:
                        appendStyleProperty(style, "mso-wrap-distance-top", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                }
            }
        }

        private static void appendStyleProperty(StringBuilder b, string propName, string propValue)
        {
            b.Append(propName);
            b.Append(":");
            b.Append(propValue);
            b.Append(";");
        }


        private static string getTextboxAnchor(uint anchor)
        {
            switch (anchor)
            {
                case 0:
                    //msoanchorTop
                    return "top";
                case 1:
                    //msoanchorMiddle
                    return "middle";
                case 2:
                    //msoanchorBottom
                    return "bottom";
                case 3:
                    //msoanchorTopCentered
                    return "top-center";
                case 4:
                    //msoanchorMiddleCentered
                    return "middle-center";
                case 5:
                    //msoanchorBottomCentered
                    return "bottom-center";
                case 6:
                    //msoanchorTopBaseline
                    return "top-baseline";
                case 7:
                    //msoanchorBottomBaseline
                    return "bottom-baseline";
                case 8:
                    //msoanchorTopCenteredBaseline
                    return "top-center-baseline";
                case 9:
                    //msoanchorBottomCenteredBaseline
                    return "bottom-center-baseline";
                default:
                    return "top";
            }
        }

        private static string mapVerticalPosition(ShapeOptions.PositionVertical vPos)
        {
            switch (vPos)
            {
                case ShapeOptions.PositionVertical.msopvAbs:
                    return "absolute";
                case ShapeOptions.PositionVertical.msopvTop:
                    return "top";
                case ShapeOptions.PositionVertical.msopvCenter:
                    return "center";
                case ShapeOptions.PositionVertical.msopvBottom:
                    return "bottom";
                case ShapeOptions.PositionVertical.msopvInside:
                    return "inside";
                case ShapeOptions.PositionVertical.msopvOutside:
                    return "outside";
                default:
                    return "absolute";
            }
        }

        private static string mapVerticalPositionRelative(ShapeOptions.PositionVerticalRelative vRel)
        {
            switch (vRel)
            {
                case ShapeOptions.PositionVerticalRelative.msoprvMargin:
                    return "margin";
                case ShapeOptions.PositionVerticalRelative.msoprvPage:
                    return "page";
                case ShapeOptions.PositionVerticalRelative.msoprvText:
                    return "text";
                case ShapeOptions.PositionVerticalRelative.msoprvLine:
                    return "line";
                default:
                    return "margin";
            }
        }

        private static string mapHorizontalPosition(ShapeOptions.PositionHorizontal hPos)
        {
            switch (hPos)
            {
                case ShapeOptions.PositionHorizontal.msophAbs:
                    return "absolute";
                case ShapeOptions.PositionHorizontal.msophLeft:
                    return "left";
                case ShapeOptions.PositionHorizontal.msophCenter:
                    return "center";
                case ShapeOptions.PositionHorizontal.msophRight:
                    return "right";
                case ShapeOptions.PositionHorizontal.msophInside:
                    return "inside";
                case ShapeOptions.PositionHorizontal.msophOutside:
                    return "outside";
                default:
                    return "absolute";
            }
        }

        private static string mapHorizontalPositionRelative(ShapeOptions.PositionHorizontalRelative hRel)
        {
            switch (hRel) 
            {
                case ShapeOptions.PositionHorizontalRelative.msoprhMargin:
                    return "margin";
                case ShapeOptions.PositionHorizontalRelative.msoprhPage:
                    return "page";
                case ShapeOptions.PositionHorizontalRelative.msoprhText:
                    return "text";
                case ShapeOptions.PositionHorizontalRelative.msoprhChar:
                    return "char";
                default:
                    return "margin";
            }
        }

        /// <summary>
        /// Generates a string id for the given shape
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private static string getShapeId(Shape shape)
        {
            StringBuilder id = new StringBuilder();
            id.Append("_x0000_s");
            id.Append(shape.spid);
            return id.ToString();
        }

    }
}
