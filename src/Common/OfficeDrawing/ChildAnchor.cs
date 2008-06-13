using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Drawing;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(0xF00F)]
    public class ChildAnchor : Record
    {
        /// <summary>
        /// Rectangle that describe sthe bounds of the anchor
        /// </summary>
        public Rectangle rcgBounds;

        public ChildAnchor(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int left = this.Reader.ReadInt32();
            int top = this.Reader.ReadInt32();
            int right = this.Reader.ReadInt32();
            int bottom = this.Reader.ReadInt32();
            this.rcgBounds = new Rectangle(
                new Point(left, top),
                new Size((right-left), (bottom-top))
            );
        }
    }
}
