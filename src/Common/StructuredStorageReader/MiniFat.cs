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

namespace DIaLOGIKa.b2xtranslator.StructuredStorageReader
{

    /// <summary>
    /// Represents the MiniFat in a compound file
    /// Author: math
    /// </summary>
    internal class MiniFat : AbstractFat
    {
        List<UInt32> _sectorsUsedByMiniFat;
        List<UInt32> _sectorsUsedByMiniStream;
        Fat _fat;
        UInt32 _miniStreamStart;
        UInt64 _sizeOfMiniStream;

        override internal UInt16 SectorSize
        {
            get { return _header.MiniSectorSize; }
        }


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fat">Handle to the Fat of the compount file</param>
        /// <param name="header">Handle to the header of the compount file</param>
        /// <param name="fileHandler">Handle to the file handler of the compount file</param>
        /// <param name="miniStreamStart">Address of the sector where the mini stream starts</param>
        internal MiniFat(Fat fat, Header header, FileHandler fileHandler, UInt32 miniStreamStart, UInt64 sizeOfMiniStream)
            : base(header, fileHandler)
        {                        
            _fat = fat;
            _miniStreamStart = miniStreamStart;
            _sizeOfMiniStream = sizeOfMiniStream;
            Init();
        }


        /// <summary>
        /// Seeks to a given position in a sector of the mini stream
        /// </summary>
        /// <param name="sector">The sector to seek to</param>
        /// <param name="position">The position in the sector to seek to</param>
        /// <returns></returns>
        override internal long SeekToPositionInSector(long sector, long position)
        {
            int sectorInMiniStreamChain = (int)((sector * _header.MiniSectorSize) / _fat.SectorSize);
            int offsetInSector = (int)((sector * _header.MiniSectorSize) % _fat.SectorSize);     

            if (position < 0)
            {
                throw new ArgumentOutOfRangeException("position");
            }

            return _fileHandler.SeekToPositionInSector(_sectorsUsedByMiniStream[sectorInMiniStreamChain], offsetInSector + position);
        }


        /// <summary>
        /// Returns the next sector in a chain
        /// </summary>
        /// <param name="currentSector">The current sector in the chain</param>
        /// <returns>The next sector in the chain</returns>
        override protected UInt32 GetNextSectorInChain(UInt32 currentSector)
        {
            UInt32 sectorInFile = _sectorsUsedByMiniFat[(int)(currentSector / _addressesPerSector)];
            // calculation of position:
            // currentSector % _addressesPerSector = number of address in the sector address
            // address uses 32 bit = 4 bytes
            _fileHandler.SeekToPositionInSector(sectorInFile, 4 * ((int)currentSector % _addressesPerSector));
            return _fileHandler.ReadUInt32();
        }


        /// <summary>
        /// Initalizes the Fat
        /// </summary>
        private void Init()
        {
            ReadSectorsUsedByMiniFAT();
            ReadSectorsUsedByMiniStream();
            CheckConsistency();
        }


        /// <summary>
        /// Reads the sectors used by the MiniFat
        /// </summary>
        private void ReadSectorsUsedByMiniFAT()
        {
            if (_header.MiniFatStartSector == SectorId.ENDOFCHAIN || _header.NoSectorsInMiniFatChain == 0x0)
            {
                return;
            }
            _sectorsUsedByMiniFat = _fat.GetSectorChain(_header.MiniFatStartSector, _header.NoSectorsInMiniFatChain, "MiniFat");
        }

        /// <summary>
        /// Reads the sectors used by the MiniFat
        /// </summary>
        private void ReadSectorsUsedByMiniStream()
        {
            if (_miniStreamStart == SectorId.ENDOFCHAIN)
            {
                return;
            }
            _sectorsUsedByMiniStream = _fat.GetSectorChain(_miniStreamStart, (UInt64)Math.Ceiling((double)_sizeOfMiniStream / _header.SectorSize), "MiniStream");
        }


        /// <summary>
        /// Checks whether the size specified in the header matches the actual size
        /// </summary>
        private void CheckConsistency()
        {
            if (_sectorsUsedByMiniFat.Count != _header.NoSectorsInMiniFatChain)
            {
                throw new ChainSizeMismatchException("MiniFat");
            }
            if (_sectorsUsedByMiniStream.Count != Math.Ceiling((double)_sizeOfMiniStream / _header.SectorSize))
            {
                throw new ChainSizeMismatchException("MiniStream");
            }

        }
    }
}
