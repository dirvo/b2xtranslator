﻿using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AxisParentSequence : BiffRecordSequence
    {
        public AxisParent AxisParent;
        public Begin Begin;
        public Pos Pos;
        public AxesSequence AxesSequence;
        public List<CrtSequence> CrtSequences;
        public End End;

        public AxisParentSequence(IStreamReader reader)
            : base(reader)
        {
            // AXISPARENT = AxisParent Begin Pos [AXES] 1*4CRT End

            // AxisParent
            this.AxisParent = (AxisParent)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // Pos
            this.Pos = (Pos)BiffRecord.ReadRecord(reader);

            // [AXES]
            RecordType next = BiffRecord.GetNextRecordType(reader);
            if (next == RecordType.Axis || next == RecordType.Text || next == RecordType.PlotArea)
            {
                this.AxesSequence = new AxesSequence(reader);
            }

            // 1*4CRT
            this.CrtSequences = new List<CrtSequence>();
            while(BiffRecord.GetNextRecordType(reader) == RecordType.ChartFormat)
            {
                this.CrtSequences.Add(new CrtSequence(reader));
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
