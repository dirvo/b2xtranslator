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

namespace ExcelprocessingMLMapping
{
    public class FillStyleMapping
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
                case 0x0008: return "00FFFF";
                case 0x0009: return "00FFFF";
                case 0x000A: return "00FFFF";
                case 0x000B: return "00FFFF";
                case 0x000C: return "00FFFF";
                case 0x000D: return "00FFFF";
                case 0x000E: return "00FFFF";
                case 0x000F: return "00FFFF";

                case 0x0010: return "00FFFF";
                case 0x0011: return "00FFFF";
                case 0x0012: return "00FFFF";
                case 0x0013: return "00FFFF";
                case 0x0014: return "00FFFF";
                case 0x0015: return "00FFFF";
                case 0x0016: return "00FFFF";
                case 0x0017: return "00FFFF";
                case 0x0018: return "00FFFF";
                case 0x0019: return "00FFFF";
                case 0x001A: return "00FFFF";
                case 0x001B: return "00FFFF";
                case 0x001C: return "00FFFF";
                case 0x001D: return "00FFFF";
                case 0x001E: return "00FFFF";
                case 0x001F: return "CCCCFF";

                case 0x0020: return "00FFFF";
                case 0x0021: return "00FFFF";
                case 0x0022: return "00FFFF";
                case 0x0023: return "00FFFF";
                case 0x0024: return "00FFFF";
                case 0x0025: return "00FFFF";
                case 0x0026: return "00FFFF";
                case 0x0027: return "00FFFF";
                case 0x0028: return "00FFFF";
                case 0x0029: return "00FFFF";
                case 0x002A: return "00FFFF";
                case 0x002B: return "00FFFF";
                case 0x002C: return "00FFFF";
                case 0x002D: return "00FFFF";
                case 0x002E: return "00FFFF";
                case 0x002F: return "00FFFF";

                case 0x0031: return "00FFFF";
                case 0x0032: return "00FFFF";
                case 0x0033: return "00FFFF";
                case 0x0034: return "00FFFF";
                case 0x0035: return "00FFFF";
                case 0x0036: return "00FFFF";
                case 0x0037: return "00FFFF";
                case 0x0038: return "00FFFF";
                case 0x0039: return "00FFFF";
                case 0x003A: return "00FFFF";
                case 0x003B: return "00FFFF";
                case 0x003C: return "00FFFF";
                case 0x003D: return "00FFFF";
                case 0x003E: return "00FFFF";
                case 0x003F: return "00FFFF";

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

    }

}
