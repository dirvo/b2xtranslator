using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorageReader
{
    public class VirtualStreamReader : BinaryReader, IStreamReader
    {
        public VirtualStreamReader(VirtualStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// The number of bytes to read is determined by the length of the array.
        /// </summary>
        /// <param name="buffer">Array which will contain the read bytes after successful execution.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the length of the array if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        public int Read(byte[] buffer)
        {
            return BaseStream.Read(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// </summary>
        /// <param name="buffer">Array which will contain the read bytes after successful execution.</param>
        /// <param name="count">Number of bytes to read.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        public int Read(byte[] buffer, int count)
        {
            return BaseStream.Read(buffer, 0, count);
        }

        public int ReadAtPos(byte[] buffer, long position, int count)
        {
            BaseStream.Seek(position, SeekOrigin.Begin);
            return BaseStream.Read(buffer, 0, count);
        }

        public byte[] ReadBytesAtPos(long position, int count)
        {
            BaseStream.Seek(position, SeekOrigin.Begin);
            return ReadBytes(count);
        }
    }
}
