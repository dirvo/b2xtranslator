using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FontFamilyName
    {
        /// <summary>
        /// When true, font is a TrueType font
        /// </summary>
        public bool fTrueType;

        /// <summary>
        /// Font family id
        /// </summary>
        public byte ff;

        /// <summary>
        /// Base weight of font
        /// </summary>
        public Int16 wWeight;

        /// <summary>
        /// Character set identifier
        /// </summary>
        public byte chs;

        /// <summary>
        /// Name of font
        /// </summary>
        public String xszFtn;

        /// <summary>
        /// Parses the byte to retrieve a FFN structure
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public FontFamilyName(byte[] bytes)
        {
            if (bytes.Length > 40)
            {
                this.wWeight = System.BitConverter.ToInt16(bytes, 2);
                this.chs = bytes[4];

                //copy name to array
                byte[] name = new byte[bytes.Length - 40];
                Array.Copy(bytes, 40, name, 0, name.Length);
                this.xszFtn = Encoding.Unicode.GetString(name);

                //replace zero termination
                this.xszFtn = this.xszFtn.Replace("\0", "");
            }
        }
    }
}
