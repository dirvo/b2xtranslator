using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TableCellPropertiesMapping : 
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private int _cellIndex;
        private XmlElement _tcPr;

        public TableCellPropertiesMapping(XmlWriter writer, int cellIndex)
            : base(writer)
        {
            _tcPr = _nodeFactory.CreateElement("w", "tcPr", OpenXmlNamespaces.WordprocessingML);
            _cellIndex = cellIndex;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            foreach (SinglePropertyModifier sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)
	            {
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
