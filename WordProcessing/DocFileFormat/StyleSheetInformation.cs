using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class StyleSheetInformation
    {
        public struct LSD
        {
            UInt16 grflsd;
            bool fLocked;
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
	    /// size of each lsd in mpstilsd.<br/>
        /// The count of lsd's is stiMaxWhenSaved
	    /// </summary>
        public UInt16 cbLSD;

        /// <summary>
        /// latent style data (size == stiMaxWhenSaved upon save!)
        /// </summary>
	    public LSD[] mpstilsd;	

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
            
            //this.rgftcStandardChpStsh = new UInt16[3];
            //this.rgftcStandardChpStsh[0] = System.BitConverter.ToUInt16(bytes, 12);
            //this.rgftcStandardChpStsh[1] = System.BitConverter.ToUInt16(bytes, 14);
            //this.rgftcStandardChpStsh[2] = System.BitConverter.ToUInt16(bytes, 16);
            //this.mpstilsd = new LSD[this.stiMaxWhenSaved];
        }
    }
}
