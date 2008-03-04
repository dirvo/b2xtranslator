/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */


using System;
using System.Collections.Generic;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorageReader
{

    /// <summary>
    /// Provides methods for accessing the file stream
    /// Author: math
    /// </summary>
    internal class FileHandler
    {
        FileStream _fileStream;
        InternalBitConverter _bitConverter;
        Header _header;
        
        /// <summary>
        /// Constructor, opens the given file
        /// </summary>        
        public FileHandler(string fileName)
        {
            _fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        }


        /// <summary>
        /// Closes the file associated with this handler
        /// </summary>
        public void CloseFile()
        {
            if (_fileStream != null)
            {
                _fileStream.Close();
            }
        }


        /// <summary>
        /// The size of the associated stream in bytes
        /// </summary>
        internal UInt64 FileStreamSize
        {
            get { return (UInt64)_fileStream.Length; }
        }


        /// <summary>
        /// Initializes the internal bit converter
        /// </summary>
        /// <param name="isLittleEndian">flag whether big endian or little endian is used</param>
        internal void InitBitConverter(bool isLittleEndian)
        {
            _bitConverter = new InternalBitConverter(isLittleEndian);            
        }


        /// <summary>
        /// Initializes the reference to the header
        /// </summary>
        /// <param name="header"></param>
        internal void SetHeaderReference(Header header)
        {
            _header = header;                
        }


        /// <summary>
        /// Seeks relative to the current position by the given offset
        /// </summary>
        internal long RelativeSeek(long offset)
        {
            if (offset < 0)
            {
                throw new ArgumentOutOfRangeException("offset");
            }
            return _fileStream.Seek(offset, SeekOrigin.Current);
        }
        

        /// <summary>
        /// Seeks to a given sector in the compound file.
        /// May only be used after SetHeaderReference() is called.
        /// </summary>
        internal long SeekToSector(long sector)
        {
            if (_header == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            if (sector < 0)
            {
                throw new ArgumentOutOfRangeException("sector");
            }
            
            // header sector == -1
            if (sector == -1)
            {
                return _fileStream.Seek(0, SeekOrigin.Begin);
            }
            return _fileStream.Seek((sector << _header.SectorShift) + Measures.HeaderSize, SeekOrigin.Begin);
        }


        /// <summary>
        /// Seeks to a given sector and position in the compound file.
        /// May only be used after SetHeaderReference() is called.
        /// </summary>
        /// <returns>The new position in the stream.</returns>
        internal long SeekToPositionInSector(long sector, long position)
        {
            if (_header == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            if (position < 0 || position >= _header.SectorSize)
            {
                throw new ArgumentOutOfRangeException("position");
            }

            // header sector == -1
            if (sector == -1)
            {
                return _fileStream.Seek(position, SeekOrigin.Begin);
            }
            return _fileStream.Seek((sector << _header.SectorShift) + Measures.HeaderSize + position, SeekOrigin.Begin);
        }


        /// <summary>
        /// Reads a byte at the current position of the file stream.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The byte value read from the stream. </returns>
        internal byte ReadByte()
        {
            int result = _fileStream.ReadByte(); 
            if (result == -1)
            {
                throw new ReadBytesAmountMismatchException();
            }
            return (byte)result;
        }


        /// <summary>
        /// Reads bytes at the current position of the file stream into a byte array.
        /// The array size determines the number of bytes to read.
        /// Advances the stream pointer accordingly.
        /// </summary>
        internal void Read(byte[] array)
        {
            Read(array, 0, array.Length);
        }


        /// <summary>
        /// Reads bytes at the current position of the file stream into a byte array.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <param name="array">The array to read to</param>
        /// <param name="offset">The offset in the array to read to</param>
        /// <param name="count">The number of bytes to read</param>        
        internal void Read(byte[] array, int offset, int count)
        {
            int result = _fileStream.Read(array, offset, count);
            if (result != count)
            {
                throw new ReadBytesAmountMismatchException();
            }
        }


        /// <summary>
        /// Reads a byte at the current position of the file stream.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The byte cast to an int, or -1 if reading from the end of the stream.</returns>
        internal int UncheckedReadByte()
        {
            return _fileStream.ReadByte();            
        }


        /// <summary>
        /// Reads bytes at the current position of the file stream into a byte array.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <param name="array">The array to read to</param>
        /// <param name="offset">The offset in the array to read to</param>
        /// <param name="count">The number of bytes to read</param>    
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        internal int UncheckedRead(byte[] array, int offset, int count)
        {
            return _fileStream.Read(array, offset, count);            
        }


        /// <summary>
        /// Reads bytes at the given position of the file stream into a byte array.
        /// The array size determines the number of bytes to read.
        /// Advances the stream pointer accordingly.
        /// </summary>
        internal void ReadPosition(byte[] array, long position)
        {
            if (position < 0)
            {
                throw new ArgumentOutOfRangeException("position");
            }
            _fileStream.Seek(position, 0);
            int result = _fileStream.Read(array, 0, array.Length);
            if (result != array.Length)
            {
                throw new ReadBytesAmountMismatchException();
            }
        }


        /// <summary>
        /// Reads a UInt16 at the current position of the file stream.
        /// May only be used after InitBitConverter() is called.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The UInt16 value read from the stream.</returns>
        internal UInt16 ReadUInt16()
        {
            if (_bitConverter == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            byte[] array = new byte[2];
            Read(array);
            return _bitConverter.ToUInt16(array);
        }


        /// <summary>
        /// Reads a UInt32 at the current position of the file stream.
        /// May only be used after InitBitConverter() is called.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The UInt32 value read from the stream.</returns>
        internal UInt32 ReadUInt32()
        {
            if (_bitConverter == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            byte[] array = new byte[4];
            Read(array);
            return _bitConverter.ToUInt32(array);
        }


        /// <summary>
        /// Reads a UInt64 at the current position of the file stream.
        /// May only be used after InitBitConverter() is called.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The UInt64 value read from the stream.</returns>
        internal UInt64 ReadUInt64()
        {
            if (_bitConverter == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            byte[] array = new byte[8];
            Read(array);
            return _bitConverter.ToUInt64(array);
        }


        /// <summary>
        /// Reads a UInt16 at the given position of the file stream.
        /// May only be used after InitBitConverter() is called.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The UInt16 value read at the given position.</returns>
        internal UInt16 ReadUInt16(long position)
        {
            if (_bitConverter == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            if (position < 0)
            {
                throw new ArgumentOutOfRangeException("position");
            }
            byte[] array = new byte[2];
            ReadPosition(array, position);
            return _bitConverter.ToUInt16(array);
        }


        /// <summary>
        /// Reads a UInt32 at the given position of the file stream.
        /// May only be used after InitBitConverter() is called.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The UInt32 value read at the given position.</returns>
        internal UInt32 ReadUInt32(long position)
        {
            if (_bitConverter == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            if (position < 0)
            {
                throw new ArgumentOutOfRangeException("position");
            }
            byte[] array = new byte[4];
            ReadPosition(array, position);
            return _bitConverter.ToUInt32(array);
        }


        /// <summary>
        /// Reads a UInt64 at the given position of the file stream.
        /// May only be used after InitBitConverter() is called.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <returns>The UInt64 value read at the given position.</returns>
        internal UInt64 ReadUInt64(long position)
        {
            if (_bitConverter == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            if (position < 0)
            {
                throw new ArgumentOutOfRangeException("position");
            }
            byte[] array = new byte[8];
            ReadPosition(array, position);
            return _bitConverter.ToUInt64(array);
        }


        /// <summary>
        /// Reads a UTF-16 encoded unicode string at the current position of the file stream.
        /// May only be used after InitBitConverter() is called.
        /// Advances the stream pointer accordingly.
        /// </summary>
        /// <param name="size">The maximum size of the string in bytes (1 char = 2 bytes) including the Unicode NULL.</param>
        /// <returns>The string read from the stream.</returns>
        internal string ReadString(int size)
        {
            if (_bitConverter == null)
            {
                throw new FileHandlerNotCorrectlyInitializedException();
            }

            if (size < 1)
            {
                throw new ArgumentOutOfRangeException("size");
            }
            byte[] array = new byte[size];
            Read(array);
            return _bitConverter.ToString(array);
        }  
    }
}
