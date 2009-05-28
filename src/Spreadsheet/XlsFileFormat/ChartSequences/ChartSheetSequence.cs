using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartSheetSequence : BiffRecordSequence
    {
        public BOF BOF;

        public ChartSheetContentSequence ChartSheetContentSequence;

        public ChartSheetSequence(IStreamReader reader) : base(reader)
        {
            //BOF 
            this.BOF = (BOF)BiffRecord.ReadRecord(reader);

            //CHARTSHEETCONTENT
            this.ChartSheetContentSequence = new ChartSheetContentSequence(reader);
        }
    }
}