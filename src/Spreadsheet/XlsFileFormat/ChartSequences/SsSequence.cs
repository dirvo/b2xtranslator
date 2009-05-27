using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeGraph;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class SsSequence : OfficeGraphBiffRecordSequence
    {
        public DataFormat DataFormat;

        public Begin Begin;

        public Chart3DBarShape Chart3DBarShape;

        public LineFormat LineFormat1;

        public AreaFormat AreaFormat1;

        public PieFormat PieFormat;

        public SerFmt SerFmt;

        public LineFormat LineFormat2;

        public AreaFormat AreaFormat2;

        public GelFrameSequence GelFrameSequence;

        public MarkerFormat MarkerFormat;

        public AttachedLabel AttachedLabel;

        public End End;

        public SsSequence(IStreamReader reader)
            : base(reader)
        {
            // SS = DataFormat Begin [Chart3DBarShape] [LineFormat AreaFormat PieFormat] 
            // [SerFmt] [LineFormat] [AreaFormat] [GELFRAME] [MarkerFormat] [AttachedLabel] End

            // DataFormat
            this.DataFormat = (DataFormat)OfficeGraphBiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)OfficeGraphBiffRecord.ReadRecord(reader);

            // [Chart3DBarShape]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                    DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.Chart3DBarShape)
            {
                this.Chart3DBarShape = (Chart3DBarShape)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // [LineFormat AreaFormat PieFormat] 
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.LineFormat)
            {
                this.LineFormat1 = (LineFormat)OfficeGraphBiffRecord.ReadRecord(reader);
                if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                    DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.AreaFormat)
                {
                    this.AreaFormat1 = (AreaFormat)OfficeGraphBiffRecord.ReadRecord(reader);
                    if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                        DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.PieFormat)
                    {
                        this.PieFormat = (PieFormat)OfficeGraphBiffRecord.ReadRecord(reader);
                    }
                }
            }

            // this is for the case that LineFormat and AreaFormat 
            // exists and is behind the SerFmt which doesn't exists 
            if (this.PieFormat == null)
            {
                if (this.LineFormat1 != null)
                {
                    this.LineFormat2 = this.LineFormat1;
                    this.LineFormat1 = null;
                }
                if (this.AreaFormat1 != null)
                {
                    this.AreaFormat2 = this.AreaFormat1;
                    this.AreaFormat1 = null;
                }
            }

            // [SerFmt]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                    DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.SerFmt)
            {
                this.SerFmt = (SerFmt)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // [LineFormat] [AreaFormat] [GELFRAME] [MarkerFormat] [AttachedLabel] End

            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.LineFormat)
            {
                this.LineFormat2 = (LineFormat)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // [AreaFormat]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.AreaFormat)
            {
                this.AreaFormat2 = (AreaFormat)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // [GELFRAME]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.GelFrame)
            {
                this.GelFrameSequence = new GelFrameSequence(reader) ;
            }

            // [MarkerFormat]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.MarkerFormat)
            {
                this.MarkerFormat = (MarkerFormat)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // [AttachedLabel]
            if (OfficeGraphBiffRecord.GetNextRecordNumber(reader) ==
                DIaLOGIKa.b2xtranslator.OfficeGraph.GraphRecordNumber.AttachedLabel)
            {
                this.AttachedLabel = (AttachedLabel)OfficeGraphBiffRecord.ReadRecord(reader);
            }

            // End 
            this.End = (End)OfficeGraphBiffRecord.ReadRecord(reader);
        }
    }
}
