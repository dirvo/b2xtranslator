using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using System.Xml;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class MainDocumentMapping : DocumentMapping
    {
        public MainDocumentMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart)
        {
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            //start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "document", OpenXmlNamespaces.WordprocessingML);

            //write namespaces
            _writer.WriteAttributeString("xmlns", "v", null, OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("xmlns", "o", null, OpenXmlNamespaces.Office);
            _writer.WriteAttributeString("xmlns", "w10", null, OpenXmlNamespaces.OfficeWord);
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            _writer.WriteStartElement("w", "body", OpenXmlNamespaces.WordprocessingML);

            //convert the document
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];
            Int32 cp = 0;
            while (cp < doc.FIB.ccpText)
            {
                Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
                ParagraphPropertyExceptions papx = findValidPapx(fc);
                TableInfo tai = new TableInfo(papx);

                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    cp = writeTable(cp, tai.iTap);
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
                }
            }

            //write the section properties of the body with the last SEPX
            SectionPropertyExceptions lastSepx = _doc.SectionTable.grpsepx[_doc.SectionTable.grpsepx.Length - 1];
            lastSepx.Convert(new SectionPropertiesMapping(_writer, _ctx, _sectionNr));

            //end the document
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
