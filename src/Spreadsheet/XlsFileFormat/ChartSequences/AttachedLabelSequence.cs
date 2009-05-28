using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AttachedLabelSequence : BiffRecordSequence
    {
        public Text Text;

        public Begin Begin;

        public Pos Pos;

        public FontX FontX;

        public AlRuns AlRuns;

        public AiSequence AiSequence;

        public Frame Frame;

        public ObjectLink ObjectLink;

        public DataLabExtContents DataLabExtContents;

        public AttachedLabelSequence(IStreamReader reader)
            : base(reader)
        {
            //ATTACHEDLABEL = Text Begin Pos [FontX] [AlRuns] AI [FRAME] [ObjectLink] [DataLabExtContents] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End
        }
    }
}
