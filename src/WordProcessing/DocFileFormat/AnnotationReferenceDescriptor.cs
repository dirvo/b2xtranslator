using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AnnotationReferenceDescriptor : ByteStructure
    {
        /// <summary>
        /// The initials of the user who left the annotation.
        /// </summary>
        public string UserInitials;

        /// <summary>
        /// An index into the string table of comment author names.
        /// </summary>
        public UInt16 AuthorIndex;

        /// <summary>
        /// Identifies a bookmark.
        /// </summary>
        public Int32 BookmarkId;

        public AnnotationReferenceDescriptor(VirtualStreamReader reader) : base(reader)
        {
            //read the user initials (LPXCharBuffer9)
            Int16 cch = _reader.ReadInt16( );
            byte[] chars = _reader.ReadBytes(18);
            this.UserInitials = Encoding.Unicode.GetString(chars, 0, cch * 2);

            this.AuthorIndex = _reader.ReadUInt16();

            //skip 4 bytes
            _reader.ReadBytes(4);

            this.BookmarkId = _reader.ReadInt32();
        }
    }
}
