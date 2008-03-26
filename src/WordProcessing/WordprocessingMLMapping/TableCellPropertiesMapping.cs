using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Collections;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TableCellPropertiesMapping : 
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private int _cellIndex;
        private XmlElement _tcPr;
        private XmlElement _tcMar;

        public TableCellPropertiesMapping(XmlWriter writer, int cellIndex)
            : base(writer)
        {
            _tcPr = _nodeFactory.CreateElement("w", "tcPr", OpenXmlNamespaces.WordprocessingML);
            _tcMar = _nodeFactory.CreateElement("w", "tcMar", OpenXmlNamespaces.WordprocessingML);
            _cellIndex = cellIndex;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            foreach (SinglePropertyModifier sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)
	            {
                    //width
                    case 0xD608:
                        Int16 boundary2 = System.BitConverter.ToInt16(sprm.Arguments, 1 + ((_cellIndex + 1) * 2));
                        Int16 boundary1 = System.BitConverter.ToInt16(sprm.Arguments, 1 + (_cellIndex * 2));
                        appendDxaElement(_tcPr, "tcW", "" + (boundary2 - boundary1));
                        break;
                    
                    //margins
                    case 0xd632:
                        byte first = sprm.Arguments[0];
                        byte lim = sprm.Arguments[1];
                        byte ftsMargin = sprm.Arguments[3];
                        Int16 wMargin = System.BitConverter.ToInt16(sprm.Arguments, 4);
                        if (_cellIndex >= first && _cellIndex < lim)
                        {
                            BitArray borderBits = new BitArray(new byte[] { sprm.Arguments[2] });
                            if (borderBits[0] == true)
                                appendDxaElement(_tcMar, "top", wMargin.ToString());
                            if (borderBits[1] == true)
                                appendDxaElement(_tcMar, "left", wMargin.ToString());
                            if (borderBits[2] == true)
                                appendDxaElement(_tcMar, "bottom", wMargin.ToString());
                            if (borderBits[3] == true)
                                appendDxaElement(_tcMar, "right", wMargin.ToString());
                        }
                        break;

                    //vertical alignment
                    case 0xD62C:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        break;

                    //shading
                    case 0xD612:
                        //cell shading for cells 0-20
                        apppendCellShading(sprm.Arguments, _cellIndex);
                        break;
                    case 0xD616:
                        //cell shading for cells 21-42
                        apppendCellShading(sprm.Arguments, _cellIndex - 21);
                        break;
                    case 0xD60C:
                        //cell shading for cells 43-62
                        apppendCellShading(sprm.Arguments, _cellIndex - 43);
                        break;
	            }
            }

            //append margins
            if (_tcMar.ChildNodes.Count > 0)
            {
                _tcPr.AppendChild(_tcMar);
            }

            //write Properties
            if (_tcPr.ChildNodes.Count > 0 || _tcPr.Attributes.Count > 0)
            {
                _tcPr.WriteTo(_writer);
            }
        }

        private void apppendCellShading(byte[] sprmArg, int cellIndex)
        {
            byte[] shdBytes = new byte[10];
            Array.Copy(sprmArg, cellIndex*10, shdBytes, 0, shdBytes.Length);
            
            ShadingDescriptor shd = new ShadingDescriptor(shdBytes);
            appendShading(_tcPr, shd);
        }
    }
}
