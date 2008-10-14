using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class Converter
    {
        public static void Convert(WordDocument doc, string outputFilename)
        {
            ConversionContext context = new ConversionContext(doc);
            WordprocessingDocument docx = null;

            //detect the document type
            if (doc.FIB.fDot)
            {
                //template
                if (doc.CommandTable.MacroDatas != null)
                {
                    //macro enabled template
                    docx = WordprocessingDocument.Create( getOutputFilename(outputFilename, WordprocessingDocumentType.MacroEnabledTemplate), 
                        WordprocessingDocumentType.MacroEnabledTemplate);
                }
                else
                {
                    //without macros
                    docx = WordprocessingDocument.Create(getOutputFilename(outputFilename, WordprocessingDocumentType.Template), 
                        WordprocessingDocumentType.Template);
                }
            }
            else
            {
                //no template
                if (doc.CommandTable.MacroDatas != null)
                {
                    //macro enabled document
                    docx = WordprocessingDocument.Create(getOutputFilename(outputFilename, WordprocessingDocumentType.MacroEnabledDocument), 
                        WordprocessingDocumentType.MacroEnabledDocument);
                }
                else
                {
                    docx = WordprocessingDocument.Create(getOutputFilename(outputFilename, WordprocessingDocumentType.Document), 
                        WordprocessingDocumentType.Document);
                }
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
                if (docx.DocumentType == WordprocessingDocumentType.MacroEnabledDocument ||
                    docx.DocumentType == WordprocessingDocumentType.MacroEnabledTemplate)
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
    
        private static string getOutputFilename(string inputfilename, WordprocessingDocumentType doctype)
        {
            string outExt = ".docx";
            switch (doctype)
            {
                case WordprocessingDocumentType.Document:
                    outExt = ".docx";
                    break;
                case WordprocessingDocumentType.MacroEnabledDocument:
                    outExt = ".docm";
                    break;
                case WordprocessingDocumentType.MacroEnabledTemplate:
                    outExt = ".dotm";
                    break;
                case WordprocessingDocumentType.Template:
                    outExt = ".dotx";
                    break;
                default:
                    outExt = ".docx";
                    break;
            }

            string inExt = Path.GetExtension(inputfilename);
            if (inExt != null)
            {
                return inputfilename.Replace(inExt, outExt);
            }
            else
            {
                return inputfilename + outExt;
            }
        }
    }
}
