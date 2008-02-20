using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class BorderCode
    {
        /// <summary>
        /// 24-bit border color
        /// </summary>
        public Int32 cv;

        /// <summary>
        /// Width of a single line in 1/8pt, max of 32pt
        /// </summary>
        public Int32 dptLineWidth;

        /// <summary>
        /// Border type code:
        /// 0 none
        /// 1 single
        /// 2 thick
        /// 3 double
        /// 5 hairline
        /// 6 dot
        /// 7 dash large gap
        /// 8 dot dash
        /// 9 dot dot dash
        /// 10 triple
        /// 11 thin-thick small gap
        /// 12 tick-thin small gap
        /// 13 thin-thick-thin small gap
        /// 14 thin-thick medium gap
        /// 15 thick-thin medium gap
        /// 16 thin-thick-thin medium gap
        /// 17 thin-thick large gap
        /// 18 thick-thin large gap
        /// 19 thin-thick-thin large gap
        /// 20 wave
        /// 21 double wave
        /// 22 dash small gap
        /// 23 dash dot stroked
        /// 24 emboss 3D
        /// 25 engrave 3D
        /// </summary>
        public Int32 brcType;

        /// <summary>
        /// Width of space to maintain between border and text within border
        /// </summary>
        public Int32 dptSpace;

        /// <summary>
        /// When true, border is drawn with shadow. Must be false when BRC is substructure of the TC
        /// </summary>
        public bool fShadow;

        /// <summary>
        /// When true, don't reverse the border
        /// </summary>
        public bool fFrame;

        /// <summary>
        /// Creates a new BorderCode with default values
        /// </summary>
        public BorderCode()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the byte for a BRC
        /// </summary>
        /// <param name="bytes"></param>
        public BorderCode(byte[] bytes)
        {
            if (bytes.Length == 8)
            {
                this.cv = System.BitConverter.ToInt32(bytes, 0);

                Int32 val = System.BitConverter.ToInt32(bytes, 4);

                this.dptLineWidth = val & 0x000000FF;
                this.brcType = val & 0x0000FF00;
                this.dptSpace = val & 0x001F0000;
                this.fShadow = Utils.BitmaskToBool(val, 0x00200000);
                this.fFrame = Utils.BitmaskToBool(val, 0x00400000);
            }
            else
            {
                throw new ByteParseException("BRC");
            }
        }

        private void setDefaultValues()
        {
            this.brcType = 0;
            this.cv = 0;
            this.dptLineWidth = 0;
            this.dptSpace = 0;
            this.fFrame = false;
            this.fShadow = false;
        }
    }
}
