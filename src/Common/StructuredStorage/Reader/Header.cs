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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Encapsulates the header of a compound file
    /// Author: math
    /// </summary>
    internal class Header : AbstractHeader
    {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fileHandler">The Handle to the file handler of the compound file</param>
        internal Header(InputHandler fileHandler)
        {
            _ioHandler = fileHandler;
            _ioHandler.SetHeaderReference(this);
            ReadHeader();
        }


        /// <summary>
        /// Reads the header from the file stream
        /// </summary>
        private void ReadHeader()
        {
            InputHandler fileHandler = ((InputHandler)_ioHandler);

            // Determine endian
            byte[] byteArray16 = new byte[2];
            fileHandler.ReadPosition(byteArray16, 0x1C);
            if (byteArray16[0] == 0xFF && byteArray16[1] == 0xFE)
            {
                fileHandler.InitBitConverter(false);
            }
            else
            {
                // default little endian
                fileHandler.InitBitConverter(true);
            }

            ulong magicNumber = fileHandler.ReadUInt64(0x0);
            // Check for Magic Number                       
            if (magicNumber != MAGIC_NUMBER)
            {                
                throw new MagicNumberException(String.Format("Found: {0,10:X}", magicNumber));
            }

            SectorShift = fileHandler.ReadUInt16(0x1E);
            MiniSectorShift = fileHandler.ReadUInt16();
            
            NoSectorsInDirectoryChain4KB = fileHandler.ReadUInt32(0x28);
            NoSectorsInFatChain = fileHandler.ReadUInt32();
            DirectoryStartSector = fileHandler.ReadUInt32();                      
            
            MiniSectorCutoff = fileHandler.ReadUInt32(0x38);
            MiniFatStartSector = fileHandler.ReadUInt32();
            NoSectorsInMiniFatChain = fileHandler.ReadUInt32();
            DiFatStartSector = fileHandler.ReadUInt32();
            NoSectorsInDiFatChain = fileHandler.ReadUInt32(); 
        }
    }
}
