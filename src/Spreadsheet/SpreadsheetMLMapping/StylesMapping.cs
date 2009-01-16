/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;
using ExcelprocessingMLMapping;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class StylesMapping : AbstractOpenXmlMapping,
          IMapping<StyleData>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public StylesMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddStylesPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Styles xml document 
        /// </summary>
        /// <param name="sd">StyleData Object</param>
        public void Apply(StyleData sd)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("styleSheet", OpenXmlNamespaces.StylesML);


            // Format mapping 
            _writer.WriteStartElement("numFmts");
            _writer.WriteAttributeString("count", sd.FormatDataList.Count.ToString());
            foreach (FormatData format in sd.FormatDataList)
            {
                _writer.WriteStartElement("numFmt");
                _writer.WriteAttributeString("numFmtId", format.ifmt.ToString());
                _writer.WriteAttributeString("formatCode", format.formatString);
                _writer.WriteEndElement();
            }
            _writer.WriteEndElement(); 

             


            /// Font Mapping
            //<fonts count="1">
            //<font>
            //<sz val="10"/>
            //<name val="Arial"/>
            //</font>
            //</fonts>
            _writer.WriteStartElement("fonts");
            _writer.WriteAttributeString("count", sd.FontDataList.Count.ToString());
            foreach (FontData font in sd.FontDataList)
            {
                // begin font element 
                _writer.WriteStartElement("font");
                // font size 
                _writer.WriteStartElement("sz");
                _writer.WriteAttributeString("val", Convert.ToString(font.size.ToPoints(), CultureInfo.GetCultureInfo("en-US")));
                _writer.WriteEndElement();
                // font name 
                _writer.WriteStartElement("name");
                _writer.WriteAttributeString("val", font.fontName);
                _writer.WriteEndElement();
                // font family 
                if (font.fontFamily != 0)
                {
                    _writer.WriteStartElement("family");
                    _writer.WriteAttributeString("val", font.fontFamily.ToString());
                    _writer.WriteEndElement();
                }
                // font charset 
                if (font.charSet != 0)
                {
                    _writer.WriteStartElement("charset");
                    _writer.WriteAttributeString("val", font.charSet.ToString());
                    _writer.WriteEndElement();
                }

                // bool values 
                if (font.isBold)
                    _writer.WriteElementString("b", "");

                if (font.isItalic)
                    _writer.WriteElementString("i", "");

                if (font.isOutline)
                    _writer.WriteElementString("outline", "");

                if (font.isShadow)
                    _writer.WriteElementString("shadow", "");

                if (font.isStrike)
                    _writer.WriteElementString("strike", "");

                // underline style mapping 
                if (font.uStyle != UnderlineStyle.none)
                {
                    _writer.WriteStartElement("u");
                    if (font.uStyle == UnderlineStyle.singleLine)
                    {
                        _writer.WriteAttributeString("val", "single");
                    }
                    else if (font.uStyle == UnderlineStyle.doubleLine)
                    {
                        _writer.WriteAttributeString("val", "double");
                    }
                    else
                    {
                        _writer.WriteAttributeString("val", font.uStyle.ToString());
                    }
                    _writer.WriteEndElement();
                }

                if (font.vertAlign != SuperSubScriptStyle.none)
                {
                    _writer.WriteStartElement("vertAlign");
                    _writer.WriteAttributeString("val", font.vertAlign.ToString());
                    _writer.WriteEndElement();
                }

                // colormapping 
                string color = StyleMappingHelper.convertColorIdToRGB(font.color);
                if (color.Equals("Auto"))
                {
                    _writer.WriteStartElement("color");
                    _writer.WriteAttributeString("auto", "1");
                    _writer.WriteEndElement();
                }
                else if (color.Equals(""))
                {
                    // do nothing 
                }
                else
                {
                    // <bgColor rgb="FFFFFF00"/>
                    _writer.WriteStartElement("color");
                    _writer.WriteAttributeString("rgb", "FF" + color);
                    _writer.WriteEndElement();
                }
 

                // end font element 
                _writer.WriteEndElement();
            }
            // write fonts end element 
            _writer.WriteEndElement(); 

            /// Fill Mapping 
            //<fills count="2">
            //<fill>
            //<patternFill patternType="none"/>
            //</fill>           
            _writer.WriteStartElement("fills");
            _writer.WriteAttributeString("count", sd.FillDataList.Count.ToString());
            foreach (FillData fd in sd.FillDataList)
            {
                _writer.WriteStartElement("fill");
                _writer.WriteStartElement("patternFill");
                _writer.WriteAttributeString("patternType", StyleMappingHelper.getStringFromFillPatern(fd.Fillpatern));
               // foreground color 
                string foregroundclr = StyleMappingHelper.convertColorIdToRGB(fd.IcvFore); 
                if (foregroundclr.Equals("Auto"))
                {
                    _writer.WriteStartElement("fgColor");
                    _writer.WriteAttributeString("auto", "1");
                    _writer.WriteEndElement();
                }
                else if (foregroundclr.Equals(""))
                {
                    // do nothing 
                }
                else
                {
                    // <fgColor rgb="FFFFFF00"/>
                    _writer.WriteStartElement("fgColor");
                    _writer.WriteAttributeString("rgb", "FF" + foregroundclr);
                    _writer.WriteEndElement();
                }
                // background color 
                string backgroundclr = StyleMappingHelper.convertColorIdToRGB(fd.IcvBack);
                if (backgroundclr.Equals("Auto"))
                {
                    _writer.WriteStartElement("bgColor");
                    _writer.WriteAttributeString("auto", "1");
                    _writer.WriteEndElement();
                }
                else if (backgroundclr.Equals(""))
                {
                    // do nothing 
                }
                else
                {
                    // <bgColor rgb="FFFFFF00"/>
                    _writer.WriteStartElement("bgColor");
                    _writer.WriteAttributeString("rgb", "FF"+backgroundclr);
                    _writer.WriteEndElement();
                }
                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();


            /// Border Mapping 
            //<borders count="1">
            //  <border>
            //      <left/>
            //      <right/>
            //      <top/>
            //      <bottom/>
            //      <diagonal/>
            //  </border>
            //</borders>
            _writer.WriteStartElement("borders");
            _writer.WriteAttributeString("count", sd.BorderDataList.Count.ToString());
            foreach (BorderData borderData in sd.BorderDataList)
            {
                _writer.WriteStartElement("border");

                // write diagonal settings 
                if (borderData.diagonalValue == 1)
                {
                    _writer.WriteAttributeString("diagonalDown", "1");
                }
                else if (borderData.diagonalValue == 2)
                {
                    _writer.WriteAttributeString("diagonalUp", "1");
                }
                else if (borderData.diagonalValue == 3)
                {
                    _writer.WriteAttributeString("diagonalDown", "1");
                    _writer.WriteAttributeString("diagonalUp", "1");
                }
                else
                {
                    // do nothing !
                }

                string borderColor = "";
                string borderStyle = ""; 

                // left border 
                _writer.WriteStartElement("left");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.left.style); 
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    borderColor = StyleMappingHelper.convertColorIdToRGB(borderData.left.colorId);
                    if (borderColor.Equals("Auto"))
                    {
                        _writer.WriteStartElement("bgColor");
                        _writer.WriteAttributeString("auto", "1");
                        _writer.WriteEndElement();
                    }
                    else if (borderColor.Equals(""))
                    {
                        // do nothing 
                    }
                    else
                    {
                        // <bgColor rgb="FFFFFF00"/>
                        _writer.WriteStartElement("color");
                        _writer.WriteAttributeString("rgb", "FF" + borderColor);
                        _writer.WriteEndElement();
                    }
                }
                _writer.WriteEndElement();

                // right border 
                _writer.WriteStartElement("right");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.right.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    borderColor = StyleMappingHelper.convertColorIdToRGB(borderData.right.colorId);
                    if (borderColor.Equals("Auto"))
                    {
                        _writer.WriteStartElement("bgColor");
                        _writer.WriteAttributeString("auto", "1");
                        _writer.WriteEndElement();
                    }
                    else if (borderColor.Equals(""))
                    {
                        // do nothing 
                    }
                    else
                    {
                        // <bgColor rgb="FFFFFF00"/>
                        _writer.WriteStartElement("color");
                        _writer.WriteAttributeString("rgb", "FF" + borderColor);
                        _writer.WriteEndElement();
                    }
                }
                _writer.WriteEndElement();

                // top border 
                _writer.WriteStartElement("top");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.top.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    borderColor = StyleMappingHelper.convertColorIdToRGB(borderData.top.colorId);
                    if (borderColor.Equals("Auto"))
                    {
                        _writer.WriteStartElement("bgColor");
                        _writer.WriteAttributeString("auto", "1");
                        _writer.WriteEndElement();
                    }
                    else if (borderColor.Equals(""))
                    {
                        // do nothing 
                    }
                    else
                    {
                        // <bgColor rgb="FFFFFF00"/>
                        _writer.WriteStartElement("color");
                        _writer.WriteAttributeString("rgb", "FF" + borderColor);
                        _writer.WriteEndElement();
                    }
                }
                _writer.WriteEndElement();

                // bottom border 
                _writer.WriteStartElement("bottom");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.bottom.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    borderColor = StyleMappingHelper.convertColorIdToRGB(borderData.bottom.colorId);
                    if (borderColor.Equals("Auto"))
                    {
                        _writer.WriteStartElement("bgColor");
                        _writer.WriteAttributeString("auto", "1");
                        _writer.WriteEndElement();
                    }
                    else if (borderColor.Equals(""))
                    {
                        // do nothing 
                    }
                    else
                    {
                        // <bgColor rgb="FFFFFF00"/>
                        _writer.WriteStartElement("color");
                        _writer.WriteAttributeString("rgb", "FF" + borderColor);
                        _writer.WriteEndElement();
                    }
                }
                _writer.WriteEndElement();

                // diagonal border 
                _writer.WriteStartElement("diagonal");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.diagonal.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    borderColor = StyleMappingHelper.convertColorIdToRGB(borderData.diagonal.colorId);
                    if (borderColor.Equals("Auto"))
                    {
                        _writer.WriteStartElement("bgColor");
                        _writer.WriteAttributeString("auto", "1");
                        _writer.WriteEndElement();
                    }
                    else if (borderColor.Equals(""))
                    {
                        // do nothing 
                    }
                    else
                    {
                        // <bgColor rgb="FFFFFF00"/>
                        _writer.WriteStartElement("color");
                        _writer.WriteAttributeString("rgb", "FF" + borderColor);
                        _writer.WriteEndElement();
                    }
                }
                // end diagonal
                _writer.WriteEndElement();
                // end border 
                _writer.WriteEndElement();
            }
            // end borders 
            _writer.WriteEndElement(); 

            ///<cellStyleXfs count="1">
            ///<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            ///</cellStyleXfs> 
            // xfcellstyle mapping 
            _writer.WriteStartElement("cellStyleXfs");
            _writer.WriteAttributeString("count", sd.XFCellStyleDataList.Count.ToString());
            foreach (XFData xfcellstyle in sd.XFCellStyleDataList)
            {
                _writer.WriteStartElement("xf");
                _writer.WriteAttributeString("numFmtId", xfcellstyle.ifmt.ToString());
                _writer.WriteAttributeString("fontId", xfcellstyle.fontId.ToString());
                _writer.WriteAttributeString("fillId", xfcellstyle.fillId.ToString());
                _writer.WriteAttributeString("borderId", xfcellstyle.borderId.ToString());

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();




            ///<cellXfs count="6">
            ///<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            // xfcell mapping 
            _writer.WriteStartElement("cellXfs");
            _writer.WriteAttributeString("count", sd.XFCellDataList.Count.ToString());
            foreach (XFData xfcell in sd.XFCellDataList)
            {
                _writer.WriteStartElement("xf");
                _writer.WriteAttributeString("numFmtId", xfcell.ifmt.ToString());
                _writer.WriteAttributeString("fontId", xfcell.fontId.ToString());
                _writer.WriteAttributeString("fillId", xfcell.fillId.ToString());
                _writer.WriteAttributeString("borderId", xfcell.borderId.ToString());
                _writer.WriteAttributeString("xfId", xfcell.ixfParent.ToString());

                // applyNumberFormat="1"
                if (xfcell.ifmt != 0)
                {
                    _writer.WriteAttributeString("applyNumberFormat", "1");
                }

                // applyFill="1"
                if (xfcell.fillId != 0)
                {
                    _writer.WriteAttributeString("applyFill", "1");
                }

                // applyFont="1"
                if (xfcell.fontId != 0)
                {
                    _writer.WriteAttributeString("applyFont", "1");
                }


                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); 




            /// write cell styles 
            /// <cellStyles count="1">
            /// <cellStyle name="Normal" xfId="0" builtinId="0"/>
            /// </cellStyles>
            /// 
            _writer.WriteStartElement("cellStyles");
            //_writer.WriteAttributeString("count", sd.StyleList.Count.ToString());
            foreach (STYLE style in sd.StyleList)
            {
                _writer.WriteStartElement("cellStyle");

                if (style.rgch != null)
                { 
                    _writer.WriteAttributeString("name", style.rgch); 
                }
                // theres a bug with the zero based reading from the referenz id 
                // so the style.ixfe value is reduzed by one
                if (style.ixfe != 0)
                {
                    _writer.WriteAttributeString("xfId", (style.ixfe - 1).ToString());
                }
                else
                {
                    _writer.WriteAttributeString("xfId", (style.ixfe).ToString());
                }
                _writer.WriteAttributeString("builtinId", style.istyBuiltIn.ToString());

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); 
            
            // close tags 
            
            _writer.WriteEndElement();      // close 
            _writer.WriteEndDocument();

            // close writer 
            _writer.Flush();
        }
    }
}