using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class FooterMapping : DocumentMapping
    {
        private CharacterRange _ftr;

        public FooterMapping(ConversionContext ctx, FooterPart part, CharacterRange ftr)
            : base(ctx, part)
        {
            _ftr = ftr;
        }
        
        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "ftr", OpenXmlNamespaces.WordprocessingML);

            //convert the footer text
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];
            Int32 cp = _ftr.CharacterPosition;
            //ignore the last 1 \r chars
            while (cp < (_ftr.CharacterPosition + _ftr.CharacterCount))
            {
                Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
                ParagraphPropertyExceptions papx = findValidPapx(fc);
                TableInfo tai = new TableInfo(papx);

                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    cp = writeTable(cp);
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
