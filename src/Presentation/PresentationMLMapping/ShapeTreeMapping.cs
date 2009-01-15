/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class ShapeTreeMapping :
        AbstractOpenXmlMapping,
        IMapping<PPDrawing>
    {
        protected int _idCnt;
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

            //TraceLogger.DebugInternal(method.ToString());

            try
            {
                method.Invoke(this, new Object[] { record });
            }
            catch (TargetInvocationException e)
            {
                TraceLogger.DebugInternal(e.InnerException.ToString());
                throw e.InnerException;
            }
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
               
                System.IO.MemoryStream ms = new System.IO.MemoryStream(clientData.bytes);
                Record rec = Record.ReadRecord(ms, 0);
                OEPlaceHolderAtom placeholder = (OEPlaceHolderAtom)rec; // clientData.FirstChildWithType<OEPlaceHolderAtom>();

                if (placeholder != null)
                {

                    _writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);

                    if (!placeholder.IsObjectPlaceholder())
                    {
                        string typeValue = Utils.PlaceholderIdToXMLValue(placeholder.PlacementId);
                        _writer.WriteAttributeString("type", typeValue);
                    }

                    if (placeholder.Position != -1)
                    {
                        _writer.WriteAttributeString("idx", placeholder.Position.ToString());
                    }

                    _writer.WriteEndElement();
                }
            }

            _writer.WriteEndElement();

            _writer.WriteEndElement();


            // Visible shape properties
            _writer.WriteStartElement("p", "spPr", OpenXmlNamespaces.PresentationML);

            ClientAnchor anchor = container.FirstChildWithType<ClientAnchor>();

            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

                _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                _writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                _writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                _writer.WriteEndElement();

                _writer.WriteEndElement();
            }

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

            _writer.WriteStartElement("a", "bodyPr", OpenXmlNamespaces.DrawingML);
            // TODO...
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "lstStyle", OpenXmlNamespaces.DrawingML);
            // TODO...
            _writer.WriteEndElement();

            new TextMapping(_ctx, _writer).Apply(textbox);

            _writer.WriteEndElement();
        }


        public void Apply(RegularContainer container)
        {
            // Descend into container records by default
            foreach (Record record in container.Children)
            {
                DynamicApply(record);
            }
        }

        public void Apply(Record record)
        {
            // Ignore unsupported records
            //TraceLogger.DebugInternal("Unsupported record: {0}", record);
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
            WriteXFrm(_writer, new Rectangle()); // groupShape.rcgBounds
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
