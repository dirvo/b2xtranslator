using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class TabDescriptor
    {
        /// <summary>
        /// Justification code:<br/>
        /// 0 left tab<br/>
        /// 1 centered tab<br/>
        /// 2 right tab<br/>
        /// 3 decimal tab<br/>
        /// 4 bar
        /// </summary>
        public byte jc;

        /// <summary>
        /// Tab leader code:<br/>
        /// 0 no leader<br/>
        /// 1 dotted leader<br/>
        /// 2 hyphenated leader<br/>
        /// 3 single line leader<br/>
        /// 4 heavy line leader<br/>
        /// 5 middle dot
        /// </summary>
        public byte tlc;

        /// <summary>
        /// Parses the bytes to retrieve a TabDescriptor
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public TabDescriptor(byte[] bytes)
        {
            if (bytes.Length == 1)
            {
                this.jc = Convert.ToByte(Convert.ToInt32(bytes[0]) & 0x07);
                this.tlc = Convert.ToByte(Convert.ToInt32(bytes[0]) & 0x38);
            }
            else
            {
                throw new ByteParseException("TBD");
            }
        }
    }
}
