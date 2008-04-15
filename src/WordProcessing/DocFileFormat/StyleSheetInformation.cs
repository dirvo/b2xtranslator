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
using System.Collections;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class StyleSheetInformation
    {
        public struct LatentStyleData
        {
            public UInt32 grflsd;
            public bool fLocked;
        }

        /// <summary>
        /// Count of styles in stylesheet
        /// </summary>
        public UInt16 cstd;
	
        /// <summary>
        /// Length of STD Base as stored in a file
        /// </summary>
        public UInt16 cbSTDBaseInFile;
	
        /// <summary>
        /// Are built-in stylenames stored?
        /// </summary>
        public bool fStdStylenamesWritten;
						
        /// <summary>
        /// Max sti known when this file was written
        /// </summary>
        public UInt16 stiMaxWhenSaved;

        /// <summary>
        /// How many fixed-index istds are there?
        /// </summary>
	    public UInt16 istdMaxFixedWhenSaved;

        /// <summary>
        /// Current version of built-in stylenames
        /// </summary>
	    public UInt16 nVerBuiltInNamesWhenSaved;

        /// <summary>
        /// This is a list of the default fonts for this style sheet.<br/>
        /// The first is for ASCII characters (0-127), the second is for East Asian characters, 
        /// and the third is the default font for non-East Asian, non-ASCII text.
        /// </summary>
	    public UInt16[] rgftcStandardChpStsh;	

	    /// <summary>
	    /// Size of each lsd in mpstilsd<br/>
        /// The count of lsd's is stiMaxWhenSaved
	    /// </summary>
        public UInt16 cbLSD;

        /// <summary>
        /// latent style data (size == stiMaxWhenSaved upon save!)
        /// </summary>
	    public LatentStyleData[] mpstilsd;	

        /// <summary>
        /// Parses the bytes to retrieve a StyleSheetInformation
        /// </summary>
        /// <param name="bytes"></param>
        public StyleSheetInformation(byte[] bytes)
        {
            this.cstd = System.BitConverter.ToUInt16(bytes, 0);
            this.cbSTDBaseInFile = System.BitConverter.ToUInt16(bytes, 2);
            if(bytes[4] == 1)
            {
                this.fStdStylenamesWritten = true;
            }
            //byte 5 is spare
            this.stiMaxWhenSaved = System.BitConverter.ToUInt16(bytes, 6);
            this.istdMaxFixedWhenSaved = System.BitConverter.ToUInt16(bytes, 8);
            this.nVerBuiltInNamesWhenSaved = System.BitConverter.ToUInt16(bytes, 10);

            this.rgftcStandardChpStsh = new UInt16[4];
            this.rgftcStandardChpStsh[0] = System.BitConverter.ToUInt16(bytes, 12);
            this.rgftcStandardChpStsh[1] = System.BitConverter.ToUInt16(bytes, 14);
            this.rgftcStandardChpStsh[2] = System.BitConverter.ToUInt16(bytes, 16);
            if (bytes.Length > 18)
            {
                this.rgftcStandardChpStsh[3] = System.BitConverter.ToUInt16(bytes, 18);
            }

            //not all stylesheet contain latent styles
            if (bytes.Length > 20)
            {
                this.cbLSD = System.BitConverter.ToUInt16(bytes, 20);
                this.mpstilsd = new LatentStyleData[this.stiMaxWhenSaved];
                for (int i = 0; i < this.mpstilsd.Length; i++)
                {
                    LatentStyleData lsd = new LatentStyleData();
                    lsd.grflsd = System.BitConverter.ToUInt32(bytes, 22 + (i * cbLSD));
                    lsd.fLocked = Utils.BitmaskToBool((int)lsd.grflsd, 0x1);
                    this.mpstilsd[i] = lsd;
                }
            }
        }
    }
}
