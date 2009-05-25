using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This structure specifies formatting information for a text run.
    /// </summary>
    public class FormatRun
    {
        /// <summary>
        /// An unsigned integer that specifies the zero-based index of the first character 
        /// of the text in the TxO record that contains the text run. 
        /// 
        /// When FormatRun is used in an array, this value MUST be in strictly increasing order.
        /// </summary>
        public UInt16 ich;

        /// <summary>
        /// A FontIndex record that specifies the font. 
        /// 
        /// If ich is equal to the length of the text, this field is undefined and MUST be ignored.
        /// </summary>
        public UInt16 ifnt;

        public FormatRun()
        {
        }

        public FormatRun(UInt16 ich, UInt16 ifnt)
        {
            this.ich = ich;
            this.ifnt = ifnt;
        }

        public FormatRun(IStreamReader reader)
        {
            ich = reader.ReadUInt16();
            ifnt = reader.ReadUInt16();
        }
    }
}
