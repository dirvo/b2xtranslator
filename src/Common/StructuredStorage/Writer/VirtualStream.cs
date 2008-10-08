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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    internal class VirtualStream
    {
        AbstractFat _fat;
        Stream _stream;
        UInt16 _sectorSize;
        OutputHandler _outputHander;


        UInt32 _startSector = SectorId.FREESECT;
        public UInt32 StartSector
        {
            get { return _startSector; }
        }
        

        public UInt64 Length
        {
            get { return (UInt64)_stream.Length; }
        }


        UInt32 _sectorCount;
        public UInt32 SectorCount
        {
            get { return _sectorCount;  }
        }


        internal VirtualStream(Stream stream, AbstractFat fat, UInt16 sectorSize, OutputHandler outputHander)
        {
            _stream = stream;
            _fat = fat;
            _sectorSize = sectorSize;
            _outputHander = outputHander;
            _sectorCount = (UInt32)Math.Ceiling((double)_stream.Length / (double)_sectorSize);
        }


        internal void write()
        {
            _startSector = _fat.writeChain(SectorCount);
            BinaryReader reader = new BinaryReader(_stream);
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            while (true) {
                byte[] bytes = reader.ReadBytes((int)_sectorSize);
                _outputHander.writeSectors(bytes, _sectorSize, (byte)0x0);
                if (bytes.Length != _sectorSize)
                {
                    break;
                }
            }
        }
    }
}
