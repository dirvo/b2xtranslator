using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartSheetContentSequence: BiffRecordSequence
    {
        public WriteProtect WriteProtect;

        public SheetExt SheetExt;

        public WebPub WebPub;

        public List<HFPicture> HFPictures;

        public PageSetupSequence PageSetupSequence;

        public PrintSize PrintSize;

        public HeaderFooter HeaderFooter;

        public BackgroundSequence BackgroundSequence;

        public List<Fbi> Fbis;

        public List<Fbi2> Fbi2s;

        public ClrtClient ClrtClient;

        public ProtectionSequence ProtectionSequence;

        public Palette Palette;

        //public SxView

        public ChartSheetContentSequence(IStreamReader reader)
            : base(reader)
        {
            //CHARTSHEETCONTENT = [WriteProtect] [SheetExt] [WebPub] *HFPicture PAGESETUP PrintSize [HeaderFooter] [BACKGROUND] *Fbi *Fbi2 [ClrtClient] [PROTECTION] 
            //[Palette] [SXViewLink] [PivotChartBits] [SBaseRef] [MsoDrawingGroup] OBJECTS Units CHARTFOMATS SERIESDATA *WINDOW *CUSTOMVIEW [CodeName] [CRTMLFRT] EOF
        }
    }
}
