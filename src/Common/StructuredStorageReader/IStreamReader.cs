using System;
using System.IO;


namespace DIaLOGIKa.b2xtranslator.StructuredStorageReader
{
    public interface IStreamReader
    {
        /// <summary>
        /// Exposes access to the underlying stream of type IStreamReader.
        /// </summary>
        /// <returns>The underlying stream associated with the IStreamReader</returns>
        Stream BaseStream { get; }

        /// <summary>
        /// Closes the current reader and the underlying stream.
        /// </summary>
        void Close();
        
        /// <summary>
        /// Returns the next available character and does not advance the byte or character position.
        /// </summary>
        /// <returns>
        /// The next available character, or -1 if no more characters are available or 
        /// the stream does not support seeking.
        /// </returns>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        int PeekChar();

        /// <summary>
        /// Reads characters from the underlying stream and advances the current position
        /// of the stream in accordance with the Encoding used and the specific character
        /// being read from the stream.
        /// </summary>
        /// <returns>
        /// The next character from the input stream, or -1 if no characters are currently available.
        /// </returns>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        int Read();
        
        //
        // Summary:
        //     Reads count bytes from the stream with index as the starting point in the
        //     byte array.
        //
        // Parameters:
        //   buffer:
        //     The buffer to read data into.
        //
        //   index:
        //     The starting point in the buffer at which to begin reading into the buffer.
        //
        //   count:
        //     The number of characters to read.
        //
        // Returns:
        //     The number of characters read into buffer. This might be less than the number
        //     of bytes requested if that many bytes are not available, or it might be zero
        //     if the end of the stream is reached.
        //
        // Exceptions:
        //   System.ArgumentException:
        //     The buffer length minus index is less than count.
        //
        //   System.ArgumentNullException:
        //     buffer is null.
        //
        //   System.ArgumentOutOfRangeException:
        //     index or count is negative.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        int Read(byte[] buffer, int index, int count);
        //
        // Summary:
        //     Reads count characters from the stream with index as the starting point in
        //     the character array.
        //
        // Parameters:
        //   buffer:
        //     The buffer to read data into.
        //
        //   index:
        //     The starting point in the buffer at which to begin reading into the buffer.
        //
        //   count:
        //     The number of characters to read.
        //
        // Returns:
        //     The total number of characters read into the buffer. This might be less than
        //     the number of characters requested if that many characters are not currently
        //     available, or it might be zero if the end of the stream is reached.
        //
        // Exceptions:
        //   System.ArgumentException:
        //     The buffer length minus index is less than count.
        //
        //   System.ArgumentNullException:
        //     buffer is null.
        //
        //   System.ArgumentOutOfRangeException:
        //     index or count is negative.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        int Read(char[] buffer, int index, int count);
        
        //
        // Summary:
        //     Reads a Boolean value from the current stream and advances the current position
        //     of the stream by one byte.
        //
        // Returns:
        //     true if the byte is nonzero; otherwise, false.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        bool ReadBoolean();
        //
        // Summary:
        //     Reads the next byte from the current stream and advances the current position
        //     of the stream by one byte.
        //
        // Returns:
        //     The next byte read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        byte ReadByte();
        //
        // Summary:
        //     Reads count bytes from the current stream into a byte array and advances
        //     the current position by count bytes.
        //
        // Parameters:
        //   count:
        //     The number of bytes to read.
        //
        // Returns:
        //     A byte array containing data read from the underlying stream. This might
        //     be less than the number of bytes requested if the end of the stream is reached.
        //
        // Exceptions:
        //   System.IO.IOException:
        //     An I/O error occurs.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.ArgumentOutOfRangeException:
        //     count is negative.
        byte[] ReadBytes(int count);
        //
        // Summary:
        //     Reads the next character from the current stream and advances the current
        //     position of the stream in accordance with the Encoding used and the specific
        //     character being read from the stream.
        //
        // Returns:
        //     A character read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        //
        //   System.ArgumentException:
        //     A surrogate character was read.
        char ReadChar();
        //
        // Summary:
        //     Reads count characters from the current stream, returns the data in a character
        //     array, and advances the current position in accordance with the Encoding
        //     used and the specific character being read from the stream.
        //
        // Parameters:
        //   count:
        //     The number of characters to read.
        //
        // Returns:
        //     A character array containing data read from the underlying stream. This might
        //     be less than the number of characters requested if the end of the stream
        //     is reached.
        //
        // Exceptions:
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        //
        //   System.ArgumentOutOfRangeException:
        //     count is negative.
        char[] ReadChars(int count);
        //
        // Summary:
        //     Reads a decimal value from the current stream and advances the current position
        //     of the stream by sixteen bytes.
        //
        // Returns:
        //     A decimal value read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        decimal ReadDecimal();
        //
        // Summary:
        //     Reads an 8-byte floating point value from the current stream and advances
        //     the current position of the stream by eight bytes.
        //
        // Returns:
        //     An 8-byte floating point value read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        double ReadDouble();
        //
        // Summary:
        //     Reads a 2-byte signed integer from the current stream and advances the current
        //     position of the stream by two bytes.
        //
        // Returns:
        //     A 2-byte signed integer read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        short ReadInt16();
        //
        // Summary:
        //     Reads a 4-byte signed integer from the current stream and advances the current
        //     position of the stream by four bytes.
        //
        // Returns:
        //     A 4-byte signed integer read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        int ReadInt32();
        //
        // Summary:
        //     Reads an 8-byte signed integer from the current stream and advances the current
        //     position of the stream by eight bytes.
        //
        // Returns:
        //     An 8-byte signed integer read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        long ReadInt64();
        //
        // Summary:
        //     Reads a signed byte from this stream and advances the current position of
        //     the stream by one byte.
        //
        // Returns:
        //     A signed byte read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        sbyte ReadSByte();
        //
        // Summary:
        //     Reads a 4-byte floating point value from the current stream and advances
        //     the current position of the stream by four bytes.
        //
        // Returns:
        //     A 4-byte floating point value read from the current stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        float ReadSingle();
        //
        // Summary:
        //     Reads a string from the current stream. The string is prefixed with the length,
        //     encoded as an integer seven bits at a time.
        //
        // Returns:
        //     The string being read.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        string ReadString();
        //
        // Summary:
        //     Reads a 2-byte unsigned integer from the current stream using little-endian
        //     encoding and advances the position of the stream by two bytes.
        //
        // Returns:
        //     A 2-byte unsigned integer read from this stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        ushort ReadUInt16();
        //
        // Summary:
        //     Reads a 4-byte unsigned integer from the current stream and advances the
        //     position of the stream by four bytes.
        //
        // Returns:
        //     A 4-byte unsigned integer read from this stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        uint ReadUInt32();
        //
        // Summary:
        //     Reads an 8-byte unsigned integer from the current stream and advances the
        //     position of the stream by eight bytes.
        //
        // Returns:
        //     An 8-byte unsigned integer read from this stream.
        //
        // Exceptions:
        //   System.IO.EndOfStreamException:
        //     The end of the stream is reached.
        //
        //   System.IO.IOException:
        //     An I/O error occurs.
        //
        //   System.ObjectDisposedException:
        //     The stream is closed.
        ulong ReadUInt64();

        int Read(byte[] buffer, int count);
        int Read(byte[] buffer);
        int ReadAtPos(byte[] buffer, long position, int count);
        byte[] ReadBytesAtPos(long position, int count);
    }
}
