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
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    internal class OutputHandler : AbstractIOHandler
    {
        internal Stream BaseStream
        {
            get { return _stream; }
        }

        /// <summary>
        /// Returns UInt64.MaxValue because size of stream is not defined yet.
        /// </summary>
        override internal UInt64 IOStreamSize
        {
            get { return UInt64.MaxValue; }
        }


        internal OutputHandler(MemoryStream memoryStream)
        {
            _stream = memoryStream;
            _bitConverter = new InternalBitConverter(true);
        }
        

        internal void writeByte(byte value)
        {
            _stream.WriteByte(value);
        }
        

        internal void writeUInt16(UInt16 value)
        {
            _stream.Write(_bitConverter.getBytes(value), 0, 2);
        }


        internal void writeUInt32(UInt32 value)
        {
            _stream.Write(_bitConverter.getBytes(value), 0, 4);
        }


        internal void writeUInt64(UInt64 value)
        {
            _stream.Write(_bitConverter.getBytes(value), 0, 8);
        }


        internal void write(byte[] data)
        {
            _stream.Write(data, 0, data.Length);
        }


        internal void writeSectors(byte[] data, UInt16 sectorSize, byte padding)
        {
            uint remaining = (uint)(data.LongLength % sectorSize);
            _stream.Write(data, 0, data.Length);
            if (remaining == 0)
            {
                return;
            }
            for (uint i = 0; i < (sectorSize - remaining) ; i++)
            {
                _stream.WriteByte(padding);
            }
        }


        internal void writeSectors(byte[] data, UInt16 sectorSize, UInt32 padding)
        {
            uint remaining = (uint)(data.LongLength % sectorSize);
            _stream.Write(data, 0, data.Length);
            if (remaining == 0)
            {
                return;
            }
            if ((sectorSize - remaining) % sizeof(UInt32) != 0)
            {
                throw new InvalidSectorSizeException();
            }

            for (uint i = 0; i < ((sectorSize - remaining)/sizeof(UInt32)); i++)
            {                                
                writeUInt32(padding);
            }
        }


        internal void writeToStream(Stream stream)
        {
            const int bytesToReadAtOnce = 512;

            BinaryReader reader = new BinaryReader(BaseStream);
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            while (true)
            {
                byte[] array = reader.ReadBytes((int)bytesToReadAtOnce);
                stream.Write(array, 0, array.Length);
                if (array.Length != bytesToReadAtOnce)
                {
                    break;
                }
            }
            stream.Flush();
        }
    }
}
