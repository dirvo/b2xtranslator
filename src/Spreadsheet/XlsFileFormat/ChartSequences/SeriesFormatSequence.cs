using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class SeriesFormatSequence : OfficeGraphBiffRecordSequence
    {
        public class LegendExceptionGroup : OfficeGraphBiffRecordSequence
        {
            public LegendException LegendException;

            public Begin Begin;

            public AttachedLabelSequence AttachedLabelSequence;

            public End End;

            

            public LegendExceptionGroup(IStreamReader reader)
                : base(reader)
            {
                // *(LegendException [Begin ATTACHEDLABEL End]) 
                this.LegendException = (LegendException)OfficeGraphBiffRecord.ReadRecord(reader);

                // [Begin ATTACHEDLABEL End]
                if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                     DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.Begin)
                {
                    // Begin 
                    this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);
                    // ATTACHEDLABEL
                    this.AttachedLabelSequence = new AttachedLabelSequence(reader);

                    

                    // End
                    this.End = (End)OfficeGraphBiffRecord.ReadRecord(reader);
                }

            }

        }


        public Series Series;

        public Begin Begin;

        public List<AiSequence> AiSequence;

        public List<SsSequence> SsSequence;

        public SerToCrt SerToCrt; 

        public SerParent SerParent ;

        public SerAuxTrend SerAuxTrend;

        public SerAuxErrBar SerAuxErrBar;

        public LegendExceptionGroup LegendExceptionSequence; 

        public End End; 

        public SeriesFormatSequence(IStreamReader reader)
            : base(reader)
        {
            // SERIESFORMAT = 
            // Series Begin 4AI *SS 
            // (SerToCrt / (SerParent (SerAuxTrend / SerAuxErrBar))) 
            // *(LegendException [Begin ATTACHEDLABEL End]) 
            // End


            // Series 
            this.Series = (Series)OfficeGraphBiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // 4AI
            this.AiSequence = new List<AiSequence>(); 
            for (int i = 0; i < 4; i++)
            {
                this.AiSequence.Add(new AiSequence(reader)); 
            }

            // *SS 
            this.SsSequence = new List<SsSequence>();
            while (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.DataFormat)
            {
                this.SsSequence.Add(new SsSequence(reader));
            }

            // (SerToCrt / (SerParent (SerAuxTrend / SerAuxErrBar)))
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.SerToCrt)
            {
                this.SerToCrt = (SerToCrt)OfficeGraphBiffRecord.ReadRecord(reader);
            }
            else
            {
                this.SerParent = (SerParent)OfficeGraphBiffRecord.ReadRecord(reader);
                // (SerAuxTrend / SerAuxErrBar)
                if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                    DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.SerAuxTrend)
                {
                    this.SerAuxTrend = (SerAuxTrend)OfficeGraphBiffRecord.ReadRecord(reader);
                }
                else
                {
                    this.SerAuxErrBar = (SerAuxErrBar)OfficeGraphBiffRecord.ReadRecord(reader); 
                }
            }


            // *(LegendException [Begin ATTACHEDLABEL End])  
            this.LegendExceptionSequence = new LegendExceptionGroup(reader); 

            // End 
            this.End = (End)OfficeGraphBiffRecord.ReadRecord(reader);

        }
    }
}
