using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    public abstract class OfficeGraphBiffRecord
    {
        IStreamReader _reader;

        RecordNumber _id;
        UInt32 _length;
        long _offset;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public OfficeGraphBiffRecord(IStreamReader reader, RecordNumber id, UInt32 length)
        {
            _reader = reader;
            _offset = _reader.BaseStream.Position;

            _id = id;
            _length = length;
        }

        public RecordNumber Id
        {
            get { return _id; }
        }

        public UInt32 Length
        {
            get { return _length; }
        }

        public long Offset
        {
            get { return _offset; }
        }

        public IStreamReader Reader
        {
            get { return _reader; }
            set { this._reader = value; }
        }
    }
}
