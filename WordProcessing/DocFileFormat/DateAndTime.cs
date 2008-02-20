using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class DateAndTime
    {
        /// <summary>
        /// minutes (0-59)
        /// </summary>
        public Int16 mint;

        /// <summary>
        /// hours (0-23)
        /// </summary>
        public Int16 hr;

        /// <summary>
        /// day of month (1-31)
        /// </summary>
        public Int16 dom;

        /// <summary>
        /// month (1-12)
        /// </summary>
        public Int16 mon;

        /// <summary>
        /// year (1900-2411)-1900
        /// </summary>
        public Int16 yr;

        /// <summary>
        /// weekday<br/>
        /// 0 Sunday
        /// 1 Monday
        /// 2 Tuesday
        /// 3 Wednesday
        /// 4 Thursday
        /// 5 Friday
        /// 6 Saturday
        /// </summary>
        public Int16 wdy;

        /// <summary>
        /// Creates a new DateAndTime with default values
        /// </summary>
        public DateAndTime()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the byte sto retrieve a DateAndTime
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public DateAndTime(byte[] bytes)
        {
            if (bytes.Length == 4)
            {
                int b0 = (int)System.BitConverter.ToInt16(bytes, 0);
                this.mint = Convert.ToInt16(b0 & 0x003F);
                this.hr = Convert.ToInt16(b0 & 0x07C0);
                this.dom = Convert.ToInt16(b0 & 0xF800);
                int b2 = (int)System.BitConverter.ToInt16(bytes, 2);
                this.mon = Convert.ToInt16(b2 & 0x000F);
                this.yr = Convert.ToInt16(b2 & 0x1FF0);
                this.wdy = Convert.ToInt16(b2 & 0xE000);
            }
            else
            {
                throw new ByteParseException("DTTM");
            }
        }
        private void setDefaultValues()
        {
            this.dom = 0;
            this.hr = 0;
            this.mint = 0;
            this.mon = 0;
            this.wdy = 0;
            this.yr = 0;
        }
    }
}
