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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    internal class StructuredStorageContext
    {
        private UInt32 _sidCounter = 0x0;

        Header _header;
        internal Header Header
        {
            get { return _header; }            
        }

        Fat _fat;
        internal Fat Fat
        {
            get { return _fat; }            
        }

        MiniFat _miniFat;
        internal MiniFat MiniFat
        {
            get { return _miniFat; }            
        }

        OutputHandler _tempOutputStream;
        internal OutputHandler TempOutputStream
        {
            get { return _tempOutputStream; }            
        }

        OutputHandler _directoryStream;
        internal OutputHandler DirectoryStream
        {
            get { return _directoryStream; }
        }

        InternalBitConverter _internalBitConverter;
        internal InternalBitConverter InternalBitConverter
        {
            get { return _internalBitConverter; }
        }

        private RootDirectoryEntry _rootDirectoryEntry;
        public RootDirectoryEntry RootDirectoryEntry
        {
            get { return _rootDirectoryEntry; }
        }

        internal StructuredStorageContext()
        {
            _tempOutputStream = new OutputHandler(new MemoryStream());
            _directoryStream = new OutputHandler(new MemoryStream());
            _header = new Header(this);
            _internalBitConverter = new InternalBitConverter(true);
            _fat = new Fat(this);
            _miniFat = new MiniFat(this);
            _rootDirectoryEntry = new RootDirectoryEntry(this);
        }

        internal UInt32 getNewSid()
        {
            return ++_sidCounter;
        }
    }
}
