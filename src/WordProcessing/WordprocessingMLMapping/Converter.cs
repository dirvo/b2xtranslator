using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class Converter
    {
        public static void Convert(WordDocument doc, string outputFilename)
        {
            ConversionContext context = new ConversionContext(doc);
            WordprocessingDocument docx = null;

            if (doc.CommandTable.MacroDatas.Count > 0)
            {
                //macro enabled document
                docx = WordprocessingDocument.Create(outputFilename.Replace(".docx", ".docm"), WordprocessingDocumentType.MacroEnabledDocument);
            }
            else
            {
                //normal document
                docx = WordprocessingDocument.Create(outputFilename, WordprocessingDocumentType.Document);
            }

            using (docx)
            {
                //Setup the writer
                XmlWriterSettings xws = new XmlWriterSettings();
                xws.OmitXmlDeclaration = false;
                xws.CloseOutput = true;
                xws.Encoding = Encoding.UTF8;
                xws.ConformanceLevel = ConformanceLevel.Document;

                //Setup the context
                context.WriterSettings = xws;
                context.Docx = docx;

                //convert the macros
                if (docx.DocumentType == WordprocessingDocumentType.MacroEnabledDocument)
                {
                    doc.Convert(new MacroBinaryMapping(context));
                    doc.Convert(new MacroDataMapping(context));
                }

                //Write styles.xml
                doc.Styles.Convert(new StyleSheetMapping(context));

                //Write numbering.xml
                doc.ListTable.Convert(new NumberingMapping(context));

                //Write fontTable.xml
                doc.FontTable.Convert(new FontTableMapping(context));

                //write document.xml and the header and footers
                doc.Convert(new MainDocumentMapping(context));

                //write the footnotes
                doc.Convert(new FootnotesMapping(context));

                //write the comments
                doc.Convert(new CommentsMapping(context));

                //write settings.xml at last because of the rsid list
                doc.DocumentProperties.Convert(new SettingsMapping(context));
            }
        }
    }
}
