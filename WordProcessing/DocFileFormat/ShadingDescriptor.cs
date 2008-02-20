using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class ShadingDescriptor
    {
        /// <summary>
        /// 24-bit foreground color
        /// </summary>
        public Int32 cvFore;

        /// <summary>
        /// 24-bit background color
        /// </summary>
        public Int32 cvBack;

        /// <summary>
        /// Shading pattern
        /// </summary>
        public UInt16 ipat;

        /// <summary>
        /// Creates a new ShadingDescriptor with default values
        /// </summary>
        public ShadingDescriptor()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a ShadingDescriptor.
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ShadingDescriptor(byte[] bytes)
        {
            if (bytes.Length == 10)
            {
                this.cvFore = System.BitConverter.ToInt32(bytes, 0);
                this.cvBack = System.BitConverter.ToInt32(bytes, 4);
                this.ipat = System.BitConverter.ToUInt16(bytes, 8);
            }
            else
            {
                throw new ByteParseException("SHD");
            }
        }

        private void setDefaultValues()
        {
            this.cvBack = 0;
            this.cvFore = 0;
            this.ipat = 0;
        }
    }
}
