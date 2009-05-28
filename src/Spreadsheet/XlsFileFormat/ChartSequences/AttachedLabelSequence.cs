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

        public CrtLayout12 CrtLayout12;

        public TextPropsSequence TextPropsSequence;

        public CrtMlfrtSequence CrtMlfrtSequence;

        public End End;

        public AttachedLabelSequence(IStreamReader reader)
            : base(reader)
        {
            //ATTACHEDLABEL = Text Begin Pos [FontX] [AlRuns] AI [FRAME] [ObjectLink] [DataLabExtContents] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End

            //Text
            this.Text = (Text)BiffRecord.ReadRecord(reader);

            //Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);
            
            //Pos 
            this.Pos = (Pos)BiffRecord.ReadRecord(reader);

            //[FontX] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.FontX)
            {
                this.FontX = (FontX)BiffRecord.ReadRecord(reader);
            }            
            
            //[AlRuns] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.AlRuns)
            {
                this.AlRuns = (AlRuns)BiffRecord.ReadRecord(reader);
            }   
            
            //AI 
            this.AiSequence = new AiSequence(reader);
            
            //[FRAME] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Frame)
            {
                this.Frame = (Frame)BiffRecord.ReadRecord(reader);
            }   
            
            //[ObjectLink] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ObjectLink)
            {
                this.ObjectLink = (ObjectLink)BiffRecord.ReadRecord(reader);
            }   
            
            //[DataLabExtContents] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExtContents)
            {
                this.DataLabExtContents = (DataLabExtContents)BiffRecord.ReadRecord(reader);
            }   
            
            //[CrtLayout12] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12)
            {
                this.CrtLayout12 = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }   
            
            //[TEXTPROPS] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.RichTextStream)
            {
                this.TextPropsSequence = new TextPropsSequence(reader);
            }  
            
            //[CRTMLFRT] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }  
            
            //End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
