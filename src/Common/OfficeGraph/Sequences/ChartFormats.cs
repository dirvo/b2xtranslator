using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    public class ChartFormats : OfficeGraphBiffRecordSequence
    {
        public Chart Chart;

        public Begin Begin;

        public List<FontList> FontLists;

        public Scl Scl;

        public ChartFormats(IStreamReader reader) : base(reader)
        {
            // Chart
            this.Chart = (Chart)OfficeGraphBiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // *2 FONTLIST
            this.FontLists = new List<FontList>();
            while(true)
            {
                FontList fl = checkForFontList(reader);
                if(fl != null)
                {
                    this.FontLists.Add(fl);
                }
                else
                {
                    break;
                }
            }
            
            // Scl
            this.Scl = (Scl)OfficeGraphBiffRecord.ReadRecord(reader);
        }

        /// <summary>
        /// If the next record initializes a FontList sequence, the sequence is parsed and returned.<br/>
        /// The position of the stream is right after the FontList sequence.<br/>
        /// Otherwise the return value is null and the position of the stream didn't changed.
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private FontList checkForFontList(IStreamReader reader)
        {
            FontList result = null;

            // read next id
            RecordNumber nextRecord = (RecordNumber)reader.ReadUInt16();

            // seek back
            reader.BaseStream.Seek(-2, System.IO.SeekOrigin.Current);

            if (nextRecord == RecordNumber.FrtFontList)
            {
                // parse FontList Sequence
                result = new FontList(reader);
            }

            return result;
        }
    }
}
