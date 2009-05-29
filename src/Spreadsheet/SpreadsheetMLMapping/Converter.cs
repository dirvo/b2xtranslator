using System;
using System.IO;
using System.Text;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.Spreadsheet;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class Converter
    {
        public static OpenXmlPackage.DocumentType DetectOutputType(XlsDocument xls)
        {
            OpenXmlPackage.DocumentType returnType = OpenXmlPackage.DocumentType.Document;
            try
            {
                //ToDo: Find better way to detect macro type
                if (xls.Storage.FullNameOfAllEntries.Contains("\\_VBA_PROJECT_CUR"))
                {
                    if (xls.WorkBookData.Template)
                    {
                        returnType = OpenXmlPackage.DocumentType.MacroEnabledTemplate;
                    }
                    else
                    {
                        returnType = OpenXmlPackage.DocumentType.MacroEnabledDocument;
                    }
                }
                else
                {
                    if (xls.WorkBookData.Template)
                    {
                        returnType = OpenXmlPackage.DocumentType.Template;
                    }
                    else
                    {
                        returnType = OpenXmlPackage.DocumentType.Document;
                    }
                }
            }
            catch (Exception)
            {
            }

            return returnType;
        }

        public static string GetConformFilename(string choosenFilename, OpenXmlPackage.DocumentType outType)
        {
            string outExt = ".xlsx";
            switch (outType)
            {
                case OpenXmlPackage.DocumentType.Document:
                    outExt = ".xlsx";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    outExt = ".xlsm";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                    outExt = ".xltm";
                    break;
                case OpenXmlPackage.DocumentType.Template:
                    outExt = ".xltx";
                    break;
                default:
                    outExt = ".xlsx";
                    break;
            }

            string inExt = Path.GetExtension(choosenFilename);
            if (inExt != null)
            {
                return choosenFilename.Replace(inExt, outExt);
            }
            else
            {
                return choosenFilename + outExt;
            }
        }

        public static void Convert(XlsDocument xls, SpreadsheetDocument xlsx)
        {
            //Setup the writer
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.CloseOutput = true;
            xws.Encoding = Encoding.UTF8;
            xws.ConformanceLevel = ConformanceLevel.Document;

            ExcelContext xlsContext = new ExcelContext(xls, xws);
            xlsContext.SpreadDoc = xlsx;

            // convert the shared string table
            if (xls.WorkBookData.SstData != null)
            {
                xls.WorkBookData.SstData.Convert(new SSTMapping(xlsContext));
            }

            // create the styles.xml
            if (xls.WorkBookData.styleData != null)
            {
                xls.WorkBookData.styleData.Convert(new StylesMapping(xlsContext));
            }

            // create the sheets
            foreach (SheetData sheet in xls.WorkBookData.boundSheetDataList)
            {
                switch (sheet.boundsheetRecord.dt)
                {
                    case BoundSheet8.SheetType.Worksheet:
                        sheet.Convert(new WorksheetMapping(xlsContext));
                        break;

                    case BoundSheet8.SheetType.Chartsheet:
                        sheet.Convert(new ChartsheetMapping(xlsContext, xlsContext.SpreadDoc.WorkbookPart.AddChartsheetPart()));
                        break;

                    default:
                        sheet.emtpyWorksheet = true;
                        sheet.Convert(new WorksheetMapping(xlsContext));
                        break;
                }
            }
            int sbdnumber = 1;
            foreach (SupBookData sbd in xls.WorkBookData.supBookDataList)
            {
                if (!sbd.SelfRef)
                {
                    sbd.Number = sbdnumber;
                    sbdnumber++;
                    sbd.Convert(new ExternalLinkMapping(xlsContext));
                }
            }

            xls.WorkBookData.Convert(new WorkbookMapping(xlsContext));

            // convert the macros
            if (xlsx.DocumentType == OpenXmlPackage.DocumentType.MacroEnabledDocument ||
                xlsx.DocumentType == OpenXmlPackage.DocumentType.MacroEnabledTemplate)
            {
                xls.Convert(new MacroBinaryMapping(xlsContext));
            }
        }
    }
}
