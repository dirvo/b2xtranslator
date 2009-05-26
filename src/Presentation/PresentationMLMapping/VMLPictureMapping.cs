using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.IO;
using System.Drawing;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Globalization;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class VMLPictureMapping
        : AbstractOpenXmlMapping
    {
        ContentPart _targetPart;
        bool _olePreview;
        private ConversionContext _ctx;
        //private XmlElement _imageData = null;

        public VMLPictureMapping(VmlPart vmlPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(vmlPart.GetStream(), xws))
        {
            _targetPart = vmlPart;  //targetPart;
            //_olePreview = olePreview;
            //_ctx = ctx;
            //_writer = writer;
            //_imageData = _nodeFactory.CreateElement("v", "imageData", OpenXmlNamespaces.VectorML);
        }
        
        public void Apply(BlipStoreEntry bse, Shape shape, ShapeOptions options, int mx, int my, int dxaGoal, int dyaGoal, ConversionContext ctx, string spid)
        {
            _ctx = ctx;
            Rectangle bounds = new Rectangle();
            ImagePart imgPart = copyPicture(bse, ref bounds);
            if (imgPart != null)
            {

                _writer.WriteStartDocument();
                _writer.WriteStartElement("xml");


                _writer.WriteStartElement("o", "shapelayout", OpenXmlNamespaces.Office);
                _writer.WriteAttributeString("v", "ext", OpenXmlNamespaces.VectorML, "edit");

                _writer.WriteStartElement("o", "idmap", OpenXmlNamespaces.Office);
                _writer.WriteAttributeString("v", "ext", OpenXmlNamespaces.VectorML, "edit");
                _writer.WriteAttributeString("data", "2");
                _writer.WriteEndElement(); //idmap

                _writer.WriteEndElement(); //shapelayout


                //v:shapetype
                PictureFrameType type = new PictureFrameType();
                type.Convert(new VMLShapeTypeMapping(_ctx, _writer));

            
                
                //v:shape
                _writer.WriteStartElement("v", "shape", OpenXmlNamespaces.VectorML);
                _writer.WriteAttributeString("id", spid);
                _writer.WriteAttributeString("type", "#" + VMLShapeTypeMapping.GenerateTypeId(type));

                StringBuilder style = new StringBuilder();
                
                string widthString = Convert.ToString(bounds.Width, CultureInfo.GetCultureInfo("en-US"));
                string heightString = Convert.ToString(bounds.Height, CultureInfo.GetCultureInfo("en-US"));
                style.Append("position:absolute;");
                style.Append("left:").Append(Convert.ToString(bounds.Left, CultureInfo.GetCultureInfo("en-US"))).Append("px;");
                style.Append("top:").Append(Convert.ToString(bounds.Top, CultureInfo.GetCultureInfo("en-US"))).Append("px;");
                style.Append("width:").Append(widthString).Append("px;");
                style.Append("height:").Append(heightString).Append("px;");
                _writer.WriteAttributeString("style", style.ToString());
                               

                if (_olePreview)
                {
                    _writer.WriteAttributeString("o", "ole", OpenXmlNamespaces.Office, "");
                }

                foreach (ShapeOptions.OptionEntry entry in options.OptionsByID.Values)
                {
                    switch (entry.pid)
                    {
                        //BORDERS

                        case ShapeOptions.PropertyId.borderBottomColor:
                            RGBColor bottomColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                            _writer.WriteAttributeString("o", "borderbottomcolor", OpenXmlNamespaces.Office, "#" + bottomColor.SixDigitHexCode);
                            break;
                        case ShapeOptions.PropertyId.borderLeftColor:
                            RGBColor leftColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                            _writer.WriteAttributeString("o", "borderleftcolor", OpenXmlNamespaces.Office, "#" + leftColor.SixDigitHexCode);
                            break;
                        case ShapeOptions.PropertyId.borderRightColor:
                            RGBColor rightColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                            _writer.WriteAttributeString("o", "borderrightcolor", OpenXmlNamespaces.Office, "#" + rightColor.SixDigitHexCode);
                            break;
                        case ShapeOptions.PropertyId.borderTopColor:
                            RGBColor topColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                            _writer.WriteAttributeString("o", "bordertopcolor", OpenXmlNamespaces.Office, "#" + topColor.SixDigitHexCode);
                            break;

                        //CROPPING

                        //case ShapeOptions.PropertyId.cropFromBottom:
                        //    //cast to signed integer
                        //    int cropBottom = (int)entry.op;
                        //    appendValueAttribute(_imageData, null, "cropbottom", cropBottom + "f", null);
                        //    break;
                        //case ShapeOptions.PropertyId.cropFromLeft:
                        //    //cast to signed integer
                        //    int cropLeft = (int)entry.op;
                        //    appendValueAttribute(_imageData, null, "cropleft", cropLeft + "f", null);
                        //    break;
                        //case ShapeOptions.PropertyId.cropFromRight:
                        //    //cast to signed integer
                        //    int cropRight = (int)entry.op;
                        //    appendValueAttribute(_imageData, null, "cropright", cropRight + "f", null);
                        //    break;
                        //case ShapeOptions.PropertyId.cropFromTop:
                        //    //cast to signed integer
                        //    int cropTop = (int)entry.op;
                        //    appendValueAttribute(_imageData, null, "croptop", cropTop + "f", null);
                        //    break;
                    }
                }

                //v:imageData
                _writer.WriteStartElement("v", "imagedata", OpenXmlNamespaces.VectorML);
                _writer.WriteAttributeString("o", "relid", OpenXmlNamespaces.Office, imgPart.RelIdToString);
                _writer.WriteAttributeString("o", "title", OpenXmlNamespaces.Office, "");
                _writer.WriteEndElement(); //imagedata
                //appendValueAttribute(_imageData, "r", "id", imgPart.RelIdToString, OpenXmlNamespaces.Relationships);
                //appendValueAttribute(_imageData, "o", "title", "", OpenXmlNamespaces.Office);
                //_imageData.WriteTo(_writer);

                //borders
                ClientAnchor anch = shape.FirstAncestorWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();
                //writePictureBorder("bordertop", anch.Top);
                //writePictureBorder("borderleft", anch.Left);
                //writePictureBorder("borderbottom", anch.Bottom);
                //writePictureBorder("borderright", anch.Right);

                //close v:shape
                _writer.WriteEndElement();

                _writer.WriteEndElement(); //xml
                _writer.WriteEndDocument();
                _writer.Flush();
            }
        }

        ///// <summary>
        ///// Writes a border element
        ///// </summary>
        ///// <param name="name">The name of the element</param>
        ///// <param name="brc">The BorderCode object</param>
        //private void writePictureBorder(string name, BorderCode brc)
        //{
        //    //_writer.WriteStartElement("w10", name, OpenXmlNamespaces.OfficeWord);
        //    //_writer.WriteAttributeString("type", getBorderType(brc.brcType));
        //    //_writer.WriteAttributeString("width", brc.dptLineWidth.ToString());
        //    //_writer.WriteEndElement();
        //}


        /// <summary>
        /// Copies the picture from the binary stream to the zip archive 
        /// and creates the relationships for the image.
        /// </summary>
        /// <param name="pict">The PictureDescriptor</param>
        /// <returns>The created ImagePart</returns>
        protected ImagePart copyPicture(BlipStoreEntry bse, ref Rectangle bounds)
        {
            //create the image part
            ImagePart imgPart = null;
            if (bse != null)
            {
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
                        throw new Exception("Cannot convert picture of type " + bse.btWin32);
                }

                imgPart.TargetDirectory = "..\\media";
                Stream outStream = imgPart.GetStream();

                Record mb = _ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                //write the blip
                if (mb != null)
                {
                    switch (bse.btWin32)
                    {
                        case BlipStoreEntry.BlipType.msoblipEMF:
                        case BlipStoreEntry.BlipType.msoblipWMF:

                            //it's a meta image
                            MetafilePictBlip metaBlip = (MetafilePictBlip)mb;
                            bounds = metaBlip.m_rcBounds;

                            //meta images can be compressed
                            byte[] decompressed = metaBlip.Decrompress();
                            outStream.Write(decompressed, 0, decompressed.Length);

                            break;
                        case BlipStoreEntry.BlipType.msoblipJPEG:
                        case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                        case BlipStoreEntry.BlipType.msoblipPNG:
                        case BlipStoreEntry.BlipType.msoblipTIFF:

                            //it's a bitmap image
                            BitmapBlip bitBlip = (BitmapBlip)mb;
                            outStream.Write(bitBlip.m_pvBits, 0, bitBlip.m_pvBits.Length);

                            break;
                    }
                }
            }
            return imgPart;
        }
    }
}
