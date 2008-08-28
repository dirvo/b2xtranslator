using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AnnotationReferenceDescriptorExtra : PlexStruct
    {
        public DateAndTime Date;

        public Int32 CommentDepth;

        public Int32 ParentOffset;

        public AnnotationReferenceDescriptorExtra(VirtualStreamReader reader)
            : base(reader)
        {
            this.Date = new DateAndTime(_reader.ReadBytes(4));
            _reader.ReadBytes(2);
            this.CommentDepth = _reader.ReadInt32();
            this.ParentOffset = _reader.ReadInt32();
            Int32 flag = _reader.ReadInt32();
        }
    }
}
