using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TextboxMapping : DocumentMapping
    {
        private ClientTextbox _textbox;

        public TextboxMapping(ConversionContext ctx, ClientTextbox textbox, ContentPart targetpart, XmlWriter writer)
            : base(ctx, targetpart, writer)
        {
            _textbox = textbox;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            _writer.WriteStartElement("v", "textbox", OpenXmlNamespaces.VectorML);
            _writer.WriteStartElement("w", "txbxContent", OpenXmlNamespaces.WordprocessingML);

            Int16 index = System.BitConverter.ToInt16(_textbox.Bytes, 2);
            index--;

            Int32 cp = 0;
            Int32 cpEnd = 0;
            BreakDescriptor bkd = null;
            Int32 txtbxSubdocStart = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn + doc.FIB.ccpEdn;

            if(_targetPart.GetType() == typeof(MainDocumentPart))
            {
                cp = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[index];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[index + 1];
                bkd = (BreakDescriptor)doc.TextboxBreakPlex.Structs[index];
            }
            if(_targetPart.GetType() == typeof(HeaderPart))
            {
                txtbxSubdocStart += doc.FIB.ccpTxbx;
                cp = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[index];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[index + 1];
                bkd = (BreakDescriptor)doc.TextboxBreakPlexHeader.Structs[index];
            }

            //convert the textbox text
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];

            while (cp < cpEnd)
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

            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.Flush();
        }
    }
}
