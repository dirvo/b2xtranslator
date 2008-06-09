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
    public class VMLPictureMapping
        : PropertiesMapping,
          IMapping<PictureDescriptor>
    {
        ContentPart _targetPart;

        public VMLPictureMapping(XmlWriter writer, ContentPart targetPart)
            : base(writer)
        {
            _targetPart = targetPart;
        }

        public void Apply(PictureDescriptor pict)
        {
            ImagePart imgPart = copyPicture(pict);
            copyPicture(pict);

            Shape shape = (Shape)pict.ShapeContainer.Children[0];
            List<ShapeOptions.OptionEntry> options = pict.ShapeContainer.ExtractOptions();

            _writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

            //v:shape
            _writer.WriteStartElement("v", "shape", OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("type", "rect");

            StringBuilder style = new StringBuilder();
            TwipsValue width = new TwipsValue(pict.dxaGoal);
            TwipsValue height = new TwipsValue(pict.dyaGoal);
            style.Append("width:").Append(width.ToPoints()).Append("pt;");
            style.Append("height:").Append(height.ToPoints()).Append("pt;");
            _writer.WriteAttributeString("style", style.ToString());

            foreach (ShapeOptions.OptionEntry entry in options)
            {
                switch (entry.pid)
                {
                    case ShapeOptions.PropertyId.borderBottomColor:
                        RGBColor bottomColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        _writer.WriteAttributeString("o", "borderbottomcolor", OpenXmlNamespaces.Office, "#"+bottomColor.ThreeDigitHexCode);
                        break;
                    case ShapeOptions.PropertyId.borderLeftColor:
                        RGBColor leftColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        _writer.WriteAttributeString("o", "borderleftcolor", OpenXmlNamespaces.Office, "#" + leftColor.ThreeDigitHexCode);
                        break;
                    case ShapeOptions.PropertyId.borderRightColor:
                        RGBColor rightColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        _writer.WriteAttributeString("o", "borderrightcolor", OpenXmlNamespaces.Office, "#" + rightColor.ThreeDigitHexCode);
                        break;
                    case ShapeOptions.PropertyId.borderTopColor:
                        RGBColor topColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        _writer.WriteAttributeString("o", "bordertopcolor", OpenXmlNamespaces.Office, "#" + topColor.ThreeDigitHexCode);
                        break;

                }
            }

            //v:imageData
            _writer.WriteStartElement("v", "imageData", OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, imgPart.RelIdToString);
            _writer.WriteEndElement();

            //borders
            writePictureBorder("bordertop", pict.brcTop);
            writePictureBorder("borderleft", pict.brcLeft);
            writePictureBorder("borderbottom", pict.brcBottom);
            writePictureBorder("borderright", pict.brcRight);

            //close v:shape
            _writer.WriteEndElement();

            //close w:pict
            _writer.WriteEndElement();
        }

        /// <summary>
        /// Writes a border element
        /// </summary>
        /// <param name="name">The name of the element</param>
        /// <param name="brc">The BorderCode object</param>
        private void writePictureBorder(string name, BorderCode brc)
        {
            _writer.WriteStartElement("w10", name, OpenXmlNamespaces.OfficeWord);
            _writer.WriteAttributeString("type", getBorderType(brc.brcType));
            _writer.WriteAttributeString("width", brc.dptLineWidth.ToString());
            _writer.WriteEndElement();
        }


        /// <summary>
        /// Copies the picture from the binary stream to the zip archive 
        /// and creates the relationships for the image.
        /// </summary>
        /// <param name="pict">The PictureDescriptor</param>
        /// <returns>The created ImagePart</returns>
        protected ImagePart copyPicture(PictureDescriptor pict)
        {
            //create the image part
            ImagePart imgPart = null;

            switch (pict.BlipStoreEntry.btWin32)
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
                    throw new MappingException("Cannot convert picture of type " + pict.BlipStoreEntry.btWin32);
            }

            //write the picture
            Stream outStream = imgPart.GetStream();
            switch (pict.BlipStoreEntry.btWin32)
            {
                case BlipStoreEntry.BlipType.msoblipEMF:
                case BlipStoreEntry.BlipType.msoblipWMF:

                    //it's a meta image
                    MetafilePictBlip metaBlip = (MetafilePictBlip)pict.BlipStoreEntry.Blip;

                    //meta images can be compressed
                    byte[] decompressed = metaBlip.Decrompress();
                    outStream.Write(decompressed, 0, decompressed.Length);

                    break;
                case BlipStoreEntry.BlipType.msoblipJPEG:
                case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                case BlipStoreEntry.BlipType.msoblipPNG:
                case BlipStoreEntry.BlipType.msoblipTIFF:

                    //it's a bitmap image
                    BitmapBlip bitBlip = (BitmapBlip)pict.BlipStoreEntry.Blip;
                    outStream.Write(bitBlip.m_pvBits, 0, bitBlip.m_pvBits.Length);

                    break;
            }

            return imgPart;
        }

    }
}
