using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.OfficeGraph;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

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

    public class FontListSequence : OfficeGraphBiffRecordSequence
    {
        public FrtFontList FrtFontList;

        public StartObject StartObject;

        public List<FontFbiWrapper> Fonts;

        public EndObject EndObject;

        public FontListSequence(IStreamReader reader) : base(reader)
        {
            //FrtFontList 
            this.FrtFontList = (FrtFontList)OfficeGraphBiffRecord.ReadRecord(reader);

            //StartObject 
            this.StartObject = (StartObject)OfficeGraphBiffRecord.ReadRecord(reader);
            
            //*(Font [Fbi]) 
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) != GraphRecordNumber.EndObject)
            {
                Font font = (Font)OfficeGraphBiffRecord.ReadRecord(reader);
                Fbi fbi = null;
                if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) == GraphRecordNumber.Fbi)
                {
                    fbi = (Fbi)OfficeGraphBiffRecord.ReadRecord(reader);
                }
                this.Fonts.Add(new FontFbiWrapper(font, fbi));
            }

            //EndObject
            this.EndObject = (EndObject)OfficeGraphBiffRecord.ReadRecord(reader);
        }
    }
}
