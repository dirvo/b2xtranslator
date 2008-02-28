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
using System.Xml;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Diagnostics;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class StyleSheetMapping 
        : AbstractOpenXmlMapping,
          IMapping<StyleSheet>
    {
        public StyleSheetMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(StyleSheet sheet)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "styles", OpenXmlNamespaces.WordprocessingML);

            foreach (StyleSheetDescription style in sheet.Styles)
            {
                if (style != null)
                {
                    _writer.WriteStartElement("w", "style", OpenXmlNamespaces.WordprocessingML);

                    _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, style.stk.ToString());
                    _writer.WriteAttributeString("w", "styleId", OpenXmlNamespaces.WordprocessingML, makeStyleId(style.xstzName));
                    
                    // <w:name val="" />
                    _writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, style.xstzName);
                    _writer.WriteEndElement();

                    // <w:basedOn val="" />
                    if (style.istdBase != 4095)
                    {
                        _writer.WriteStartElement("w", "basedOn", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, makeStyleId(sheet.Styles[(int)style.istdBase].xstzName));
                        _writer.WriteEndElement();
                    }

                    // <w:next val="" />
                    _writer.WriteStartElement("w", "next", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, makeStyleId(sheet.Styles[(int)style.istdNext].xstzName));
                    _writer.WriteEndElement();

                    // <w:link val="" />
                    _writer.WriteStartElement("w", "link", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, makeStyleId(sheet.Styles[(int)style.istdLink].xstzName));
                    _writer.WriteEndElement();

                    // <w:locked/>
                    if (style.fLocked)
                    {
                        _writer.WriteElementString("w", "locked", OpenXmlNamespaces.WordprocessingML, null);
                    }

                    // <w:hidden/>
                    if (style.fHidden)
                    {
                        _writer.WriteElementString("w", "hidden", OpenXmlNamespaces.WordprocessingML, null);
                    }

                    // <w:semiHidden/>
                    if (style.fSemiHidden)
                    {
                        _writer.WriteElementString("w", "semiHidden", OpenXmlNamespaces.WordprocessingML, null);
                    }

                    //write paragraph properties
                    if (style.papx != null)
                    {
                        foreach (SinglePropertyModifier sprm in style.papx.grpprl)
                        {
                            
                        }
                    }
                    
                    //write character properties
                    if (style.chpx != null)
                    {
                        foreach (SinglePropertyModifier sprm in style.chpx.grpprl)
                        {

                        }
                    }

                    _writer.WriteEndElement();
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndDocument();
        }

        /// <summary>
        /// Generates a style id
        /// </summary>
        /// <param name="stylename">the name of the style</param>
        /// <returns></returns>
        private string makeStyleId(string stylename)
        {
            return stylename.Replace(" ", "");
        }
    }
}
