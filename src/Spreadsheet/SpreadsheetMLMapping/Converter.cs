using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.IO;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.Spreadsheet;
using System.Xml;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;

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
                    if (xls.workBookData.Template)
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
                    if (xls.workBookData.Template)
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
            using (xlsx)
            {
                //Setup the writer
                XmlWriterSettings xws = new XmlWriterSettings();
                xws.CloseOutput = true;
                xws.Encoding = Encoding.UTF8;
                xws.ConformanceLevel = ConformanceLevel.Document;

                ExcelContext xlsContext = new ExcelContext(xls, xws);
                xlsContext.SpreadDoc = xlsx;

                // Converts the sst data !!!
                if (xls.workBookData.SstData != null)
                    xls.workBookData.SstData.Convert(new SSTMapping(xlsContext));

                // creates the styles.xml
                if (xls.workBookData.styleData != null)
                    xls.workBookData.styleData.Convert(new StylesMapping(xlsContext));

                // creates the Spreadsheets
                foreach (WorkSheetData var in xls.workBookData.boundSheetDataList)
                {
                    if (var.boundsheetRecord.sheetType == BOUNDSHEET.sheetTypes.worksheet)
                    {
                        var.Convert(new WorksheetMapping(xlsContext));
                    }
                    else
                    {
                        var.emtpyWorksheet = true;
                        var.Convert(new WorksheetMapping(xlsContext));
                    }
                }
                int sbdnumber = 1;
                foreach (SupBookData sbd in xls.workBookData.supBookDataList)
                {
                    if (!sbd.SelfRef)
                    {
                        sbd.Number = sbdnumber;
                        sbdnumber++;
                        sbd.Convert(new ExternalLinkMapping(xlsContext));
                    }
                }

                xls.workBookData.Convert(new WorkbookMapping(xlsContext));

                // convert the macros
                if (xlsx.DocumentType == OpenXmlPackage.DocumentType.MacroEnabledDocument ||
                    xlsx.DocumentType == OpenXmlPackage.DocumentType.MacroEnabledTemplate)
                {
                    xls.Convert(new MacroBinaryMapping(xlsContext));
                }
            }
        }
    }
}
