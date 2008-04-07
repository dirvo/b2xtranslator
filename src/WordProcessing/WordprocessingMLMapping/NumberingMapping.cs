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

        private enum LevelJustification
        {
            left = 0,
            right,
            center
        }

        public NumberingMapping(NumberingDefinitionsPart numPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(numPart.GetStream(), xws))
        {
        }

        public void Apply(ListTable rglst)
        {
            _writer.WriteStartElement("w", "numbering", OpenXmlNamespaces.WordprocessingML);

            for (int i = 0 ; i < rglst.Count; i++)
            {
                ListData lstf = rglst[i];

                //start abstractNum
                _writer.WriteStartElement("w", "abstractNum", OpenXmlNamespaces.WordprocessingML);

                _writer.WriteAttributeString("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML, i.ToString());

                //nsid
                _writer.WriteStartElement("w", "nsid", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, String.Format("{0:X8}", lstf.lsid));
                _writer.WriteEndElement();

                //template
                _writer.WriteStartElement("w", "tmpl", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, String.Format("{0:X8}", lstf.tplc));
                _writer.WriteEndElement();

                //multiLevelType
                _writer.WriteStartElement("w", "multiLevelType", OpenXmlNamespaces.WordprocessingML);
                if (lstf.fHybrid)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "hybridMultilevel");
                else if (lstf.fSimpleList)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "singleLevel");
                else
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "multiLevel");
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

                    //Number level text
                    _writer.WriteStartElement("w", "lvlText", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, getLvlText(lvl.NumberText));
                    _writer.WriteEndElement();
                    
                    //jc
                    _writer.WriteStartElement("w", "lvlJc", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ((LevelJustification)lvl.jc).ToString());
                    _writer.WriteEndElement();

                    _writer.WriteEndElement();
                }

                //end abstractNum
                _writer.WriteEndElement();

                //start num
                _writer.WriteStartElement("w", "num", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "numId", OpenXmlNamespaces.WordprocessingML, (i+1).ToString());
                _writer.WriteStartElement("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, i.ToString());
                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
            _writer.Flush();
        }

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
                    break;
            }
        }
    }
}
