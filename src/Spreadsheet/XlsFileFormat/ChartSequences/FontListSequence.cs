using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class FontFbiWrapper
    {
        private Font _font;
        private Fbi _fbi;

        public FontFbiWrapper(Font font, Fbi fbi)
        {
            this._font = font;
            this._fbi = fbi;
        }

        public Font Font
        {
            get { return this._font; }
            set { this._font = value; }
        }

        public Fbi Fbi
        {
            get { return this._fbi; }
            set { this._fbi = value; }
        }
    }

    public class FontListSequence : BiffRecordSequence
    {
        public FrtFontList FrtFontList;

        public StartObject StartObject;

        public List<FontFbiWrapper> Fonts;

        public EndObject EndObject;

        public FontListSequence(IStreamReader reader) : base(reader)
        {
            //FrtFontList 
            this.FrtFontList = (FrtFontList)BiffRecord.ReadRecord(reader);

            //StartObject 
            this.StartObject = (StartObject)BiffRecord.ReadRecord(reader);
            
            //*(Font [Fbi]) 
            while (BiffRecord.GetNextRecordType(reader) != RecordType.EndObject)
            {
                Font font = (Font)BiffRecord.ReadRecord(reader);
                Fbi fbi = null;
                if (BiffRecord.GetNextRecordType(reader) == RecordType.Fbi)
                {
                    fbi = (Fbi)BiffRecord.ReadRecord(reader);
                }
                this.Fonts.Add(new FontFbiWrapper(font, fbi));
            }

            //EndObject
            this.EndObject = (EndObject)BiffRecord.ReadRecord(reader);
        }
    }
}
