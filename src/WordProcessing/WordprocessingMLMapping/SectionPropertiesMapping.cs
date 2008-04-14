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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class SectionPropertiesMapping :
        PropertiesMapping,
        IMapping<SectionPropertyExceptions>
    {
        private XmlElement _sectPr;

        private enum SectionType
        {
            continuous = 0,
            nextColumn,
            nextPage,
            evenPage,
            oddPage
        }

        private enum PageOrientation
        {
            landscape = 1,
            portrait
        }

        private enum DocGridType
        {
            Default,
            lines,
            linesAndChars,
            snapToChars,
        }

        /// <summary>
        /// Creates a new SectionPropertiesMapping which writes the 
        /// properties to the given writer
        /// </summary>
        /// <param name="writer">The XmlWriter</param>
        public SectionPropertiesMapping(XmlWriter writer)
            : base(writer)
        {
            _sectPr = _nodeFactory.CreateElement("w", "sectPr", OpenXmlNamespaces.WordprocessingML);
        }

        /// <summary>
        /// Creates a new SectionPropertiesMapping which appends 
        /// the properties to a given node.
        /// </summary>
        /// <param name="sectPr">The sectPr node</param>
        public SectionPropertiesMapping(XmlElement sectPr) 
            : base(null)
        {
            _nodeFactory = sectPr.OwnerDocument;
            _sectPr = sectPr;
        }

        /// <summary>
        /// Converts the given SectionPropertyExceptions
        /// </summary>
        /// <param name="sepx"></param>
        public void Apply(SectionPropertyExceptions sepx)
        {
            XmlElement pgMar = _nodeFactory.CreateElement("w", "pgMar", OpenXmlNamespaces.WordprocessingML);
            XmlElement pgSz = _nodeFactory.CreateElement("w", "pgSz", OpenXmlNamespaces.WordprocessingML);
            XmlElement docGrid = _nodeFactory.CreateElement("w", "docGrid", OpenXmlNamespaces.WordprocessingML);
            XmlElement cols = _nodeFactory.CreateElement("w", "cols", OpenXmlNamespaces.WordprocessingML);
            XmlElement pgBorders = _nodeFactory.CreateElement("w", "pgBorders", OpenXmlNamespaces.WordprocessingML);
            XmlElement paperSrc = _nodeFactory.CreateElement("w", "paperSrc", OpenXmlNamespaces.WordprocessingML);

            foreach (SinglePropertyModifier sprm in sepx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //page margins
                    case 0xB021:
                        //left margin
                        appendValueAttribute(pgMar, "left", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0xB022:
                        //right margin
                        appendValueAttribute(pgMar, "right", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x9023:
                        //top margin
                        appendValueAttribute(pgMar, "top", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x9024:
                        //bottom margin
                        appendValueAttribute(pgMar, "bottom", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0xB025:
                        //gutter margin
                        appendValueAttribute(pgMar, "gutter", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0xB017:
                        //header margin
                        appendValueAttribute(pgMar, "header", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0xB018:
                        //footer margin
                        appendValueAttribute(pgMar, "footer", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //page size and orientation
                    case 0xb01f:
                        //width
                        appendValueAttribute(pgSz, "w", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0xb020:
                        //height
                        appendValueAttribute(pgSz, "h", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x301d:
                        //orientation
                        appendValueAttribute(pgSz, "orient", ((PageOrientation)sprm.Arguments[0]).ToString());
                        break;

                    //paper source
                    case 0x5007:
                        appendValueAttribute(paperSrc, "first", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x5008:
                        appendValueAttribute(paperSrc, "other", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //page borders
                    case 0x702B:
                    case 0xD234:
                        //top
                        XmlElement topBorder = _nodeFactory.CreateElement("w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), topBorder);
                        addOrSetBorder(pgBorders, topBorder);
                        break;
                    case 0x702C:
                    case 0xD235:
                        //left
                        XmlElement leftBorder = _nodeFactory.CreateElement("w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), leftBorder);
                        addOrSetBorder(pgBorders, leftBorder);
                        break;
                    case 0x702D:
                    case 0xD236:
                        //left
                        XmlElement bottomBorder = _nodeFactory.CreateElement("w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bottomBorder);
                        addOrSetBorder(pgBorders, bottomBorder);
                        break;
                    case 0x702E:
                    case 0xD237:
                        //left
                        XmlElement rightBorder = _nodeFactory.CreateElement("w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), rightBorder);
                        addOrSetBorder(pgBorders, rightBorder);
                        break;

                    //doc grid
                    case 0x9031:
                        appendValueAttribute(docGrid, "linePitch", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x7030:
                        appendValueAttribute(docGrid, "charSpace", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case 0x5032:
                        appendValueAttribute(docGrid, "type", ((DocGridType)System.BitConverter.ToInt16(sprm.Arguments, 0)).ToString());
                        break;

                    //columns
                    case 0x500B:
                        Int32 colNum = System.BitConverter.ToInt16(sprm.Arguments,0) + 1;
                        appendValueAttribute(cols, "num", colNum.ToString());
                        break;
                    case 0x900c:
                        appendValueAttribute(cols, "space", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //bidi
                    case 0x3228:
                        appendFlagElement(_sectPr, sprm, "bidi", true);
                        break;

                    //title page
                    case 0x300A:
                        appendFlagElement(_sectPr, sprm, "titlePg", true);
                        break;

                    //text flow
                    case 0x5033:

                        break;

                    //RTL gutter
                    case 0x322A:
                        appendFlagElement(_sectPr, sprm, "rtlGutter", true);
                        break;

                    //type
                    case 0x3009:
                        appendValueElement(_sectPr, "type", ((SectionType)sprm.Arguments[0]).ToString(), true);
                        break;

                    //align
                    case 0x301A:
                        appendValueElement(_sectPr, "vAlign", sprm.Arguments[0].ToString(), true);
                        break;
                }
            }

            //append page size
            if (pgSz.Attributes.Count > 0)
            {
                _sectPr.AppendChild(pgSz);
            }

            //append borders
            if (pgBorders.ChildNodes.Count > 0)
            {
                _sectPr.AppendChild(pgBorders);
            }

            //append margin
            if (pgMar.Attributes.Count > 0)
            {
                _sectPr.AppendChild(pgMar);
            }

            //append paper info
            if (paperSrc.Attributes.Count > 0)
            {
                _sectPr.AppendChild(paperSrc);
            }

            //append columns
            if (cols.Attributes.Count > 0 || cols.ChildNodes.Count > 0)
            {
                _sectPr.AppendChild(cols);
            }

            //append doc grid
            if (docGrid.Attributes.Count > 0)
            {
                _sectPr.AppendChild(docGrid);
            }

            if (_writer != null)
            {
                //write the properties
                _sectPr.WriteTo(_writer);
            }
        }
    }
}
