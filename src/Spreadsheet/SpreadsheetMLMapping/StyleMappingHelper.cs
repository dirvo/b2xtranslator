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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;
using System.Xml;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class StyleMappingHelper
    {
        /// <summary>
        /// This is the FillPatern mapping, it is used to convert the binary fillpatern to the open xml string
        /// </summary>
        /// <param name="fp"></param>
        /// <returns></returns>
        public static string getStringFromFillPatern(StyleEnum fp)
        {
            switch (fp)
            {
                case StyleEnum.FLSNULL: return "none";
                case StyleEnum.FLSSOLID: return "solid";
                case StyleEnum.FLSMEDGRAY: return "mediumGray";
                case StyleEnum.FLSDKGRAY: return "darkGray";
                case StyleEnum.FLSLTGRAY: return "lightGray";
                case StyleEnum.FLSDKHOR: return "darkHorizontal";
                case StyleEnum.FLSDKVER: return "darkVertical";
                case StyleEnum.FLSDKDOWN: return "darkDown";
                case StyleEnum.FLSDKUP: return "darkUp";
                case StyleEnum.FLSDKGRID: return "darkGrid";
                case StyleEnum.FLSDKTRELLIS: return "darkTrellis";
                case StyleEnum.FLSLTHOR: return "lightHorizontal";
                case StyleEnum.FLSLTVER: return "lightVertical";
                case StyleEnum.FLSLTDOWN: return "lightDown";
                case StyleEnum.FLSLTUP: return "lightUp";
                case StyleEnum.FLSLTGRID: return "lightGrid";
                case StyleEnum.FLSLTTRELLIS: return "lightTrellis";
                case StyleEnum.FLSGRAY125: return "gray125";
                case StyleEnum.FLSGRAY0625: return "gray0625";

                default: return "none";
            }
        }

        /// <summary>
        /// Method converts a colorID to a RGB color value 
        /// </summary>
        /// <param name="colorID"></param>
        /// <returns></returns>
        public static string convertColorIdToRGB(int colorID)
        {
            switch (colorID)
            {
                case 0x0000: return "000000";// Black
                case 0x0001: return "FFFFFF";// White
                case 0x0002: return "FF0000";// Red
                case 0x0003: return "00FF00";// Green 
                case 0x0004: return "0000FF";// Blue
                case 0x0005: return "FFFF00";// Yellow
                case 0x0006: return "FF00FF";// Magenta
                case 0x0007: return "00FFFF";// Cyan
                case 0x0008: return "000000";
                case 0x0009: return "FFFFFF";
                case 0x000A: return "FF0000";
                case 0x000B: return "00FF00";
                case 0x000C: return "0000FF";

                case 0x000D: return "FFFF00";
                case 0x000E: return "FF00FF";
                case 0x000F: return "00FFFF";

                case 0x0010: return "800000";
                case 0x0011: return "008000";
                case 0x0012: return "000080";
                case 0x0013: return "808000";
                case 0x0014: return "800080";
                case 0x0015: return "008080";

                case 0x0016: return "C0C0C0";
                case 0x0017: return "808080";
                case 0x0018: return "9999FF";
                case 0x0019: return "993366";
                case 0x001A: return "FFFFCC";

                case 0x001B: return "CCFFFF";
                case 0x001C: return "660066";
                case 0x001D: return "FF8080";
                case 0x001E: return "0066CC";
                case 0x001F: return "CCCCFF";

                case 0x0020: return "000080";
                case 0x0021: return "FF00FF";
                case 0x0022: return "FFFF00";
                case 0x0023: return "00FFFF";
                case 0x0024: return "800080";

                case 0x0025: return "800000";
                case 0x0026: return "008080";
                case 0x0027: return "0000FF";
                case 0x0028: return "00CCFF";
                case 0x0029: return "CCFFFF";

                case 0x002A: return "CCFFCC";
                case 0x002B: return "FFFF99";
                case 0x002C: return "99CCFF";
                case 0x002D: return "FF99CC";
                case 0x002E: return "CC99FF";

                case 0x002F: return "FFCC99";
                case 0x0030: return "3366FF";
                case 0x0031: return "33CCCC";
                case 0x0032: return "99CC00";
                case 0x0033: return "FFCC00";

                case 0x0034: return "FF9900";
                case 0x0035: return "FF6600";
                case 0x0036: return "666699";
                case 0x0037: return "969696";
                case 0x0038: return "003366";

                case 0x0039: return "339966";
                case 0x003A: return "003300";
                case 0x003B: return "333300";
                case 0x003C: return "993300";
                case 0x003D: return "993366";
                case 0x003E: return "333399";
                case 0x003F: return "333333";

                case 0x0040: return "";
                case 0x0041: return "";
                case 0x004D: return "";
                case 0x004E: return "";
                case 0x004F: return "";
                case 0x0051: return "";
                case 0x7FFF: return "Auto";
                default: return "";
            }
        }

        public static string convertBorderStyle(ushort style)
        {
            switch (style)
            {
                case 0x0000: return "none";
                case 0x0001: return "thin";
                case 0x0002: return "medium";
                case 0x0003: return "dashed";
                case 0x0004: return "dotted";
                case 0x0005: return "thick";
                case 0x0006: return "double";
                case 0x0007: return "hair";
                case 0x0008: return "mediumDashed";
                case 0x0009: return "dashDot";
                case 0x000A: return "mediumDashDot";
                case 0x000B: return "dashDotDot";
                case 0x000C: return "mediumDashDotDot";
                case 0x000D: return "slantDashDot";
                default: return "none"; 
            }
        }



        public static void addFontElement(XmlWriter _writer, FontData font, FontElementType type)
        {

            // begin font element 
            if (type == FontElementType.NormalStyle)
                _writer.WriteStartElement("font");
            else if (type == FontElementType.String)
                _writer.WriteStartElement("rPr"); 

            // font size 
            _writer.WriteStartElement("sz");
            _writer.WriteAttributeString("val", Convert.ToString(font.size.ToPoints(), CultureInfo.GetCultureInfo("en-US")));
            _writer.WriteEndElement();
            // font name 
            if (type == FontElementType.NormalStyle)
                _writer.WriteStartElement("name");
            else if (type == FontElementType.String)
                _writer.WriteStartElement("rFont"); 
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
            StylesMapping.WriteRgbColor(_writer, StyleMappingHelper.convertColorIdToRGB(font.color));

            // end font element 
            _writer.WriteEndElement(); 
        }
    }

}
