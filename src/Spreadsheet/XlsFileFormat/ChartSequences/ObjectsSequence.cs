using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ObjectsSequence : BiffRecordSequence
    {
        public ObjectsSequence(IStreamReader reader)
            : base(reader)
        {
            // OBJECTS = *(MsoDrawing *Continue [OBJ / TEXTOBJECT]) [MsoDrawingSelection]

        }
    }
}
