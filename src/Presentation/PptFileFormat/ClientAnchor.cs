using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(0xF010)]
    public class ClientAnchor : Record
    {
        public int Top;
        public int Left;
        public int Right;
        public int Bottom;

        public ClientAnchor(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Top = this.Reader.ReadUInt16();
            this.Left = this.Reader.ReadUInt16();
            this.Right = this.Reader.ReadUInt16();
            this.Bottom = this.Reader.ReadUInt16();
        }

        public override string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Top = {2}, Left = {3}, Bottom = {4}, Right = {5})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.Top, this.Left, this.Bottom, this.Right);
        }
    }
}
