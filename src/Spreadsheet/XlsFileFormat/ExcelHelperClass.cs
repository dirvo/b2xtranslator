using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Assembly of some static methods 
    /// </summary>
    public class ExcelHelperClass
    {
        /// <summary>
        /// This method is used to parse a RK Recordnumber  
        /// This is a special MS Excel format that is used to store floatingpoint and integer values 
        /// The floatingpoint arithmetic is a little bit different to the IEEE standard format.  
        /// 
        /// Integer and floating point values are booth stored in a RK Record. To differ between them 
        /// it is necessary to check the first two bits from this value. 
        /// </summary>
        /// <param name="rk">Bytestream from the RK Record</param>
        /// <returns>The correct parsed number</returns>
        public static double NumFromRK(Byte[] rk)
        {
            double num = 0;
            int high = 1023;
            UInt32 number;
            number = System.BitConverter.ToUInt32(rk, 0);
            // Select which type of number 
            UInt32 type = number & 0x00000003;
            // if the last two bits are 00 or 01 then it is floating point IEEE number 
            // type 0 and type 1 expects booth the same arithmetic 
            if (type == 0 || type == 1)
            {

                UInt32 mant = 0;
                // masking the mantisse 
                mant = number & 0x000ffffc;
                // shifting the result by 2  
                mant = mant >> 2;

                UInt32 exp = 0;
                // masking the exponent 
                exp = number & 0x7ff00000;
                // shifting the exponent by 20 
                exp = exp >> 20;
                // (1 + (Mantisse / 2^18)) * 2 ^ (Exponent - 1023) 
                num = (1 + (mant / System.Math.Pow(2.0, 18.0))) * System.Math.Pow(2, (double)(exp - high));
                // now there is a sign bit too, the highest bit from the RK Record 
                UInt32 signBit = number & 0x80000000;
                // shifting the value by 31 bit 
                signBit = signBit >> 31;
                // if the signBit is 0 it is a positive number, otherwise it is a negative number 
                if (signBit == 1)
                {
                    num *= -1;
                }

            }
            // if type is 1 then it is an IEEE number * 100 
            else if (type == 1)
            {
                num *= 100;

            }
            // if type is 2 or 3 it is an integer value
            else if (type == 2 || type == 3)
            {
                // 30 bits for the integer value, 2 bits for the type identification 
                UInt32 unumber = 0;
                unumber = number & 0xfffffffc;
                // shifting the value by two 
                unumber = unumber >> 2;
                num = (double)unumber;
            }
            // if type is 3, it has to be multiplicated with 100
            else if (type == 3)
            {
                num *= 100;
            }
            return num;
        }

        /// <summary>
        /// converts the integer column value to a string like AB 
        /// excel binary format has a cap at column 256 -> IV, so there is no need to 
        /// create an almighty algorithm ;) 
        /// </summary>
        /// <returns>String</returns>
        public static String intToABCString(int colnumber, string rownumber)
        {
            String value = "";
            int remain = 0; 
            if (colnumber < 26)
            {
                value += (char)(colnumber + 65);
            }
            else if (colnumber < Math.Pow(26, 2))
            {
                remain = colnumber % 26;
                colnumber = colnumber / 26;
                value += (char)(colnumber + 64);
                value = value + (char)(remain + 65); 
            }
            else if (colnumber < Math.Pow(26, 3))
            {
                remain = colnumber % (int)Math.Pow(26, 2);
                colnumber = colnumber / (int)Math.Pow(26, 2);
                value += (char)(colnumber + 64);
                colnumber = remain; 
                remain = colnumber % 26;
                colnumber = colnumber / 26;
                value = value + (char)(colnumber + 64);
                value = value + (char)(remain + 65); 
            }
            return value + rownumber; 
        }
    }
}
