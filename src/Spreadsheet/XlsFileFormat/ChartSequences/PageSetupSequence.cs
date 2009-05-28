using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class PageSetupSequence: BiffRecordSequence
    {

        public PageSetupSequence(IStreamReader reader)
            : base(reader)
        {
            //CHARTSHEETCONTENT = [WriteProtect] [SheetExt] [WebPub] *HFPicture PAGESETUP PrintSize [HeaderFooter] [BACKGROUND] *Fbi *Fbi2 [ClrtClient] [PROTECTION] 
            //[Palette] [SXViewLink] [PivotChartBits] [SBaseRef] [MsoDrawingGroup] OBJECTS Units CHARTFOMATS SERIESDATA *WINDOW *CUSTOMVIEW [CodeName] [CRTMLFRT] EOF
        }
    }
}
