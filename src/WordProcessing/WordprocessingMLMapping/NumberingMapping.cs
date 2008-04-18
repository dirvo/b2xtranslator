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
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class NumberingMapping : AbstractOpenXmlMapping,
          IMapping<ListTable>
    {
        private ConversionContext _ctx;

        private enum LevelJustification
        {
            left = 0,
            center,
            right
        }

        public NumberingMapping(ConversionContext ctx)
            : base(XmlWriter.Create(ctx.Docx.MainDocumentPart.AddNumberingDefinitionsPart().GetStream(), ctx.WriterSettings))
        {
            _ctx = ctx;
        }

        public void Apply(ListTable rglst)
        {
            _writer.WriteStartElement("w", "numbering", OpenXmlNamespaces.WordprocessingML);

            for (int i = 0; i < rglst.Count; i++)
            {
                ListData lstf = rglst[i];

                //start abstractNum
                _writer.WriteStartElement("w", "abstractNum", OpenXmlNamespaces.WordprocessingML);

                _writer.WriteAttributeString("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML, i.ToString());

                //nsid
                _writer.WriteStartElement("w", "nsid", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, String.Format("{0:X8}", lstf.lsid));
                _writer.WriteEndElement();

                //multiLevelType
                _writer.WriteStartElement("w", "multiLevelType", OpenXmlNamespaces.WordprocessingML);
                if (lstf.fHybrid)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "hybridMultilevel");
                else if (lstf.fSimpleList)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "singleLevel");
                else
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "multilevel");
                _writer.WriteEndElement();

                //template
                _writer.WriteStartElement("w", "tmpl", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, String.Format("{0:X8}", lstf.tplc));
                _writer.WriteEndElement();

                //writes the levels
                for (int j = 0; j < lstf.rglvl.Length; j++)
                {
                    ListLevel lvl = lstf.rglvl[j];

                    _writer.WriteStartElement("w", "lvl", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "ilvl", OpenXmlNamespaces.WordprocessingML, j.ToString());

                    //starts at
                    _writer.WriteStartElement("w", "start", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, lvl.iStartAt.ToString());
                    _writer.WriteEndElement();

                    //number format
                    _writer.WriteStartElement("w", "numFmt", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, getNumberFormat(lvl.nfc));
                    _writer.WriteEndElement();

                    //suffix
                    _writer.WriteStartElement("w", "suff", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, lvl.ixchFollow.ToString());
                    _writer.WriteEndElement();

                    //style
                    Int16 styleIndex = lstf.rgistd[j];
                    if(styleIndex != ListData.ISTD_NIL)
                    {
                        _writer.WriteStartElement("w", "pStyle", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, StyleSheetMapping.MakeStyleId(_ctx.Doc.Styles.Styles[styleIndex].xstzName));
                        _writer.WriteEndElement();
                    }

                    //Number level text
                    _writer.WriteStartElement("w", "lvlText", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, getLvlText(lvl.xst));
                    _writer.WriteEndElement();
                    
                    //jc
                    _writer.WriteStartElement("w", "lvlJc", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ((LevelJustification)lvl.jc).ToString());
                    _writer.WriteEndElement();

                    //pPr
                    lvl.grpprlPapx.Convert(new ParagraphPropertiesMapping(_writer, _ctx, null));

                    //rPr
                    lvl.grpprlChpx.Convert(new CharacterPropertiesMapping(_writer, _ctx.Doc, new RevisionData(lvl.grpprlChpx)));

                    _writer.WriteEndElement();
                }

                //end abstractNum
                _writer.WriteEndElement();
            }

            //write the overrides
            for (int i = 0; i < _ctx.Doc.ListFormatOverrideTable.Count; i++)
            {
                ListFormatOverride lfo = _ctx.Doc.ListFormatOverrideTable[i];

                //start num
                _writer.WriteStartElement("w", "num", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "numId", OpenXmlNamespaces.WordprocessingML, (i+1).ToString());

                int index = findIndexbyId(rglst, lfo.lsid);

                _writer.WriteStartElement("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, index.ToString());

                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
            _writer.Flush();
        }

        private int findIndexbyId(List<ListData> list, Int32 id)
        {
            int ret = -1;
            for (int i = 0; i < list.Count; i++)
			{
                if (list[i].lsid == id)
                {
                    ret = i;
                    break;
                }
			}
            return ret;
        }

        /// <summary>
        /// Converts the number text of the binary format to the number text of OOXML.
        /// OOXML uses different placeholders for the numbers.
        /// </summary>
        /// <param name="numberText">The number text of the binary format</param>
        /// <returns></returns>
        private string getLvlText(string numberText)
        {
            string ret = numberText;

            ret = ret.Replace(new string((char)0x0000, 1), "%1");
            ret = ret.Replace(new string((char)0x0001, 1), "%2");
            ret = ret.Replace(new string((char)0x0002, 1), "%3");
            ret = ret.Replace(new string((char)0x0003, 1), "%4");
            ret = ret.Replace(new string((char)0x0004, 1), "%5");
            ret = ret.Replace(new string((char)0x0005, 1), "%6");
            ret = ret.Replace(new string((char)0x0006, 1), "%7");
            ret = ret.Replace(new string((char)0x0007, 1), "%8");
            ret = ret.Replace(new string((char)0x0008, 1), "%9");

            return ret;
        }

        /// <summary>
        /// Converts the number format code of the binary format.
        /// </summary>
        /// <param name="nfc">The number format code</param>
        /// <returns>The OOXML attribute value</returns>
        private string getNumberFormat(byte nfc)
        {
            switch (nfc)
            {
                case 0:
                    return "decimal";
                case 1:
                    return "upperRoman";
                case 2:
                    return "lowerRoman";
                case 3:
                    return "upperLetter";
                case 4:
                    return "lowerLetter";
                case 5:
                    return "ordinal";
                case 6:
                    return "cardinalText";
                case 7:
                    return "ordinalText";
                case 23:
                    return "bullet";
                //ToDO: implement rest of the number formats
                default:
                    return "decimal";
            }
        }
    }
}
