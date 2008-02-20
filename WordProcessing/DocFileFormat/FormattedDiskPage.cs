using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class FormattedDiskPage
    {
        /// <summary>
        /// The WordDocument stream
        /// </summary>
        public VirtualStream WordStream;

        /// <summary>
        /// The offset of the page in the WordDocument stream
        /// </summary>
        public Int32 Offset;

        /// <summary>
        /// Count of runs for that FKP
        /// </summary>
        public byte crun;

        /// <summary>
        /// Each value is the limit of a paragraph or run of exception text
        /// </summary>
        public Int32[] rgfc;

        /// <summary>
        /// Returns the hex dump of the FKP
        /// </summary>
        /// <returns>The hex dump of the FKP as string</returns>
        public override string ToString()
        {
            int colCount = 16;

            byte[] bytes = new byte[512];
            this.WordStream.Read(bytes, 512, this.Offset);

            return Utils.GetHashDump(bytes);
        }
    }
}
