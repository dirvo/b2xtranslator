using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartSheetSequence : BiffRecordSequence
    {
        public BOF BOF;

        public ChartFrtInfo ChartFrtInfo;

        public ChartSheetContentSequence ChartSheetContentSequence;

        public ChartSheetSequence(IStreamReader reader) : base(reader)
        {
            //BOF 
            this.BOF = (BOF)BiffRecord.ReadRecord(reader);

            // [ChartFrtInfo] (not specified)
            if (BiffRecord.GetNextRecordType(reader) == RecordType.WriteProtect)
            {
                this.ChartFrtInfo = (ChartFrtInfo)BiffRecord.ReadRecord(reader);
            }

            //CHARTSHEETCONTENT
            this.ChartSheetContentSequence = new ChartSheetContentSequence(reader);
        }
    }
}