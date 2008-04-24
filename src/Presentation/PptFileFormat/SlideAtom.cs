using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(TypeCode = 1007)]
    public class SlideAtom : Record
    {
        public SSlideLayoutAtom Layout;
        public Int32 MasterId;
        public Int32 NotesId;
        public UInt16 Flags;

        public SlideAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Layout = new SSlideLayoutAtom(this.Reader);
            this.MasterId = this.Reader.ReadInt32();
            this.NotesId = this.Reader.ReadInt32();
            this.Flags = this.Reader.ReadUInt16();
            this.Reader.ReadUInt16(); // Throw away undocumented data
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Layout = {2}\n{1}MasterId = {3}, NotesId = {4}, Flags = {5})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.Layout, this.MasterId, this.NotesId, this.Flags);
        }
    }

}
