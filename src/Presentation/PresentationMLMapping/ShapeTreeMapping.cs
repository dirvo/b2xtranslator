using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Reflection;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Drawing;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class ShapeTreeMapping :
        AbstractOpenXmlMapping,
        IMapping<PPDrawing>
    {
        protected int _idCnt;
        protected int _placeholderCnt;
        protected ConversionContext _ctx;

        public ShapeTreeMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public void DynamicApply(Record record)
        {
            // Call Apply(record) with dynamic dispatch (selection based on run-time type of record)
            MethodInfo method = this.GetType().GetMethod("Apply", new Type[] { record.GetType() });

            Console.WriteLine(method);

            method.Invoke(this, new Object[] { record });
        }

        public void Apply(PPDrawing drawing)
        {
            Apply((RegularContainer) drawing);
        }

        public void Apply(DrawingContainer drawingContainer)
        {
            GroupContainer group = drawingContainer.FirstChildWithType<GroupContainer>();
            IEnumerator<Record> iter = group.Children.GetEnumerator();
            iter.MoveNext();

            ShapeContainer header = iter.Current as ShapeContainer;
            WriteGroupShapeProperties(header);

            while (iter.MoveNext())
                DynamicApply(iter.Current);
        }

        public void Apply(ShapeContainer container)
        {
            _writer.WriteStartElement("p", "sp", OpenXmlNamespaces.PresentationML);

            _writer.WriteStartElement("p", "nvSpPr", OpenXmlNamespaces.PresentationML);

            WriteCNvPr("");

            _writer.WriteElementString("p", "cNvSpPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteStartElement("p", "nvPr", OpenXmlNamespaces.PresentationML);
            
            ClientData clientData = container.FirstChildWithType<ClientData>();

            if (clientData != null)
            {
                OEPlaceHolderAtom placeholder = clientData.FirstChildWithType<OEPlaceHolderAtom>();

                if (placeholder != null)
                {
                    string typeValue = Utils.PlaceholderIdToXMLValue(placeholder.PlaceholderId);

                    _writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);
                    _writer.WriteAttributeString("type", typeValue);
                    _writer.WriteAttributeString("idx", _placeholderCnt.ToString());
                    _writer.WriteEndElement();

                    _placeholderCnt++;
                }
            }

            _writer.WriteEndElement();

            _writer.WriteEndElement();


            _writer.WriteStartElement("p", "spPr", OpenXmlNamespaces.PresentationML);
            // TODO: Visible shape properties...?
            _writer.WriteEndElement();   

            // Descend into unsupported records
            foreach (Record record in container.Children)
            {
                DynamicApply(record);
            }

            _writer.WriteEndElement();
        }

        public void Apply(ClientTextbox textbox)
        {
            _writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);

            TextAtom text = textbox.FirstChildWithType<TextAtom>();
            StyleTextPropAtom style = textbox.FirstChildWithType<StyleTextPropAtom>();

            _writer.WriteStartElement("a", "bodyPr", OpenXmlNamespaces.DrawingML);
            // TODO...
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "lstStyle", OpenXmlNamespaces.DrawingML);
            // TODO...
            _writer.WriteEndElement();

            //if (style == null)
            {
                _writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);
                _writer.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingML);
                _writer.WriteStartElement("a", "t", OpenXmlNamespaces.DrawingML);

                _writer.WriteValue(text.Text);

                _writer.WriteEndElement();
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                // TODO...
                _writer.WriteEndElement();

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
        }


        public void Apply(RegularContainer container)
        {
            // Descend into unsupported records
            foreach (Record record in container.Children)
            {
                DynamicApply(record);
            }
        }

        public void Apply(Record record)
        {
            Console.WriteLine("Unsupported record: {0}", record);
            // Ignore unsupported records
        }

        private void WriteGroupShapeProperties(ShapeContainer header)
        {
            GroupShapeRecord groupShape = header.FirstChildWithType<GroupShapeRecord>();

            // Write non-visible Group Shape properties
            _writer.WriteStartElement("p", "nvGrpSpPr", OpenXmlNamespaces.PresentationML);

            // Non-visible Canvas Properties
            WriteCNvPr("");

            _writer.WriteElementString("p", "cNvGrpSpPr", OpenXmlNamespaces.PresentationML, "");
            _writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");

            _writer.WriteEndElement();


            // Write visible Group Shape properties
            _writer.WriteStartElement("p", "grpSpPr", OpenXmlNamespaces.PresentationML);
            WriteXFrm(_writer, groupShape.rcgBounds);
            _writer.WriteEndElement();
        }

        private void WriteCNvPr(string name)
        {
            _writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            _writer.WriteAttributeString("id", (++_idCnt).ToString());
            _writer.WriteAttributeString("name", name);
            _writer.WriteEndElement();
        }

        private void WriteXFrm(XmlWriter _writer, Rectangle rect)
        {
            _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

            // TODO: Coordinate conversion?
            _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("x", rect.X.ToString());
            _writer.WriteAttributeString("y", rect.Y.ToString());
            _writer.WriteEndElement();

            // TODO: Coordinate conversion?
            _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("cx", rect.Width.ToString());
            _writer.WriteAttributeString("cy", rect.Height.ToString());
            _writer.WriteEndElement();

            // TODO: Where do we get this from?
            _writer.WriteStartElement("a", "chOff", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("x", "0");
            _writer.WriteAttributeString("y", "0");
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "chExt", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("cx", rect.Width.ToString());
            _writer.WriteAttributeString("cy", rect.Height.ToString());
            _writer.WriteEndElement();            

            _writer.WriteEndElement();
        }
    }
}
