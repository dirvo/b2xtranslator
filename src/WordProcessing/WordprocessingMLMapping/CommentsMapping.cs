using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class CommentsMapping : DocumentMapping
    {
        public CommentsMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.CommentsPart)
        {
            _ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;
            int index = 0;

            _writer.WriteStartElement("w", "comments", OpenXmlNamespaces.WordprocessingML);

            Int32 cpStart = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr;
            Int32 cp = cpStart;
            while (cp < (cpStart + doc.FIB.ccpAtn - 2))
            {
                _writer.WriteStartElement("w", "comment", OpenXmlNamespaces.WordprocessingML);

                AnnotationReferenceDescriptor atrdPre10 = (AnnotationReferenceDescriptor)doc.AnnotationsReferencePlex.Elements[index];
                AnnotationReferenceDescriptorExtra atrdPost10 = doc.AnnotationReferenceExtraTable[index];

                _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, index.ToString());
                _writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, doc.AuthorTable.Strings[atrdPre10.AuthorIndex]);
                atrdPost10.Date.Convert(new DateMapping(_writer));
                _writer.WriteAttributeString("w", "initials", OpenXmlNamespaces.WordprocessingML, atrdPre10.UserInitials);

                cp = writeParagraph(cp);
                _writer.WriteEndElement();
                index++;
            }

            _writer.WriteEndElement();

            _writer.Flush();
        }
    }
}
