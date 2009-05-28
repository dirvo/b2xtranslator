using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ObjectsSequence : BiffRecordSequence
    {
        public MsoDrawingSelection MsoDrawingSelection;

        public List<ObjectGroup> ObjectGroups;

        public ObjectsSequence(IStreamReader reader)
            : base(reader)
        {
            // OBJECTS = *(MsoDrawing *Continue [OBJ / TEXTOBJECT]) [MsoDrawingSelection]

            // *(MsoDrawing *Continue [OBJ / TEXTOBJECT])
            this.ObjectGroups = new List<ObjectGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.MsoDrawing)
            {
                this.ObjectGroups.Add(new ObjectGroup(reader));
            }

            // [MsoDrawingSelection]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.MsoDrawingSelection)
            {
                this.MsoDrawingSelection = (MsoDrawingSelection)BiffRecord.ReadRecord(reader);
            }
        }
    }

    public class ObjectGroup
    {
        public MsoDrawing MsoDrawing;

        public List<Continue> Continues;

        public ObjSequence ObjSequence;

        public TextObjectSequence TextObjectSequence;

        public ObjectGroup(IStreamReader reader)
        {
            // MsoDrawing *Continue [OBJ / TEXTOBJECT]

            // MsoDrawing
            this.MsoDrawing = (MsoDrawing)BiffRecord.ReadRecord(reader);

            // *Continue
            this.Continues = new List<Continue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continues.Add((Continue)BiffRecord.ReadRecord(reader));
            }

            // [OBJ / TEXTOBJECT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Obj)
            {
                this.ObjSequence = new ObjSequence(reader);
            }
            else if (BiffRecord.GetNextRecordType(reader) == RecordType.TxO)
            {
                this.TextObjectSequence = new TextObjectSequence(reader);
            }
        }
    }
}
