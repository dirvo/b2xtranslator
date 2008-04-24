using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class TextAtom : Record
    {
        public string Text;

        public TextAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance, Encoding encoding)
            : base(_reader, size, typeCode, version, instance)
        {
            byte[] bytes = new byte[size];
            this.Reader.Read(bytes, 0, (int)size);

            this.Text = new String(encoding.GetChars(bytes)) + "\n";
        }

        public override string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Text = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1), this.Text);
        }
    }

}
