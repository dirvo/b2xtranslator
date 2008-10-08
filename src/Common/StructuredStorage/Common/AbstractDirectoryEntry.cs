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

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Common
{
    public class AbstractDirectoryEntry
    {
        private UInt32 _sid;
        public UInt32 Sid
        {
            get { return _sid; }
        }

        protected string _path;
        public string Path
        {
            get { return _path + Name; }
        }


        // Name
        protected string _name;
        public string Name
        {
            get { return MaskingHandler.Mask(_name); }
            protected set {
                _name = value;
                if (_name.Length >= 32)
                {
                    throw new InvalidValueInDirectoryEntryException("_ab");                    
                }
            }
        }

        UInt16 _lengthOfName;
        public UInt16 LengthOfName
        {
            get
            {
                // length of name in bytes including unicode 0;
                _lengthOfName = (UInt16)((_name.Length + 1)*2);
                return _lengthOfName;
            }            
        }



        // Type
        DirectoryEntryType _type;
        public DirectoryEntryType Type
        {
            get { return _type; }
            protected set
            {
                if ((int)value < 0 || (int)value > 5)
                {
                    throw new InvalidValueInDirectoryEntryException("_mse");
                }
                _type = value;
            }
        }


        // Color
        DirectoryEntryColor _color;
        public DirectoryEntryColor Color
        {
            get { return _color; }
            protected set
            {
                if ((int)value < 0 || (int)value > 1)
                {
                    throw new InvalidValueInDirectoryEntryException("_bflags");
                }
                _color = value;
            }
        }


        // Left sibling sid
        UInt32 _leftSiblingSid;
        public UInt32 LeftSiblingSid
        {
            get { return _leftSiblingSid; }
            protected set { _leftSiblingSid = value; }
        }


        // Right sibling sid
        UInt32 _rightSiblingSid;
        public UInt32 RightSiblingSid
        {
            get { return _rightSiblingSid; }
            protected set { _rightSiblingSid = value; }
        }


        // Child sibling sid
        UInt32 _childSiblingSid;
        public UInt32 ChildSiblingSid
        {
            get { return _childSiblingSid; }
            protected set { _childSiblingSid = value; }
        }


        //CLSID
        Guid _clsId;
        public Guid ClsId
        {
            get { return _clsId; }
            protected set { _clsId = value; }
        }


        // User flags
        UInt32 _userFlags;
        public UInt32 UserFlags
        {
            get { return _userFlags; }
            protected set { _userFlags = value; }
        }


        // Start sector
        UInt32 _startSector;
        public UInt32 StartSector
        {
            get { return _startSector; }
            protected set { _startSector = value; }
        }


        // Size of stream in bytes
        UInt64 _sizeOfStream;
        public UInt64 SizeOfStream
        {
            get { return _sizeOfStream; }
            protected set { _sizeOfStream = value; }
        }

        internal AbstractDirectoryEntry(UInt32 sid)
        {
            _sid = sid;
        }
    }
}
