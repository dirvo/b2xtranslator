using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

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
                if (type == 1)
                {
                    num /= 100;

                }
            }
            // if type is 1 then it is an IEEE number * 100 

            // if type is 2 or 3 it is an integer value
            else if (type == 2 || type == 3)
            {
                // 30 bits for the integer value, 2 bits for the type identification 
                UInt32 unumber = 0;
                unumber = number & 0xfffffffc;
                // shifting the value by two 
                unumber = unumber >> 2;
                num = (double)unumber;
                if (type == 3)
                {
                    num /= 100;
                }
            }
            // if type is 3, it has to be multiplicated with 100

            return num;
        }

        /// <summary>
        /// converts the integer column value to a string like AB 
        /// excel binary format has a cap at column 256 -> IV, so there is no need to 
        /// create an almighty algorithm ;) 
        /// </summary>
        /// <returns>String</returns>
        public static String intToABCString(int colnumber, String rownumber)
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

        /// <summary>
        /// converts the integer column value to a string like AB 
        /// excel binary format has a cap at column 256 -> IV, so there is no need to 
        /// create an almighty algorithm ;) 
        /// </summary>
        /// <returns>String</returns>
        public static String intToABCString(int colnumber, String rownumber, bool colRelative, bool rwRelative)
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
            if (!colRelative)
                value = "$" + value;

            if (!rwRelative)
                rownumber = "$" + rownumber; 

            return value + rownumber;
        }


        public static Stack<AbstractPtg> getFormulaStack(IStreamReader reader, UInt16 cce)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();
            try
            {
                for (uint i = 0; i < cce; i++)
                {
                    PtgNumber ptgtype = (PtgNumber)reader.ReadByte();

                    if ((int)ptgtype > 0x5D)
                    {
                        ptgtype -= 0x40;
                    }

                    else if ((int)ptgtype > 0x3D)
                    {
                        ptgtype -= 0x20;
                    }
                    AbstractPtg ptg = null;
                    if (ptgtype == PtgNumber.Ptg0x19Sub)
                    {
                        Ptg0x19Sub ptgtype2 = (Ptg0x19Sub)reader.ReadByte();
                        switch (ptgtype2)
                        {
                            case Ptg0x19Sub.PtgAttrSum: ptg = new PtgAttrSum(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrIf: ptg = new PtgAttrIf(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrGoto: ptg = new PtgAttrGoto(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrSemi: ptg = new PtgAttrSemi(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrChoose: ptg = new PtgAttrChoose(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrSpace: ptg = new PtgAttrSpace(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgNotDocumented: ptg = new PtgNotDocumented(reader, ptgtype2); break;
                            default: break;
                        }
                    }
                    else if (ptgtype == PtgNumber.Ptg0x18Sub)
                    {

                    }
                    else
                    {
                        switch (ptgtype)
                        {
                            case PtgNumber.PtgInt: ptg = new PtgInt(reader, ptgtype); break;
                            case PtgNumber.PtgAdd: ptg = new PtgAdd(reader, ptgtype); break;
                            case PtgNumber.PtgSub: ptg = new PtgSub(reader, ptgtype); break;
                            case PtgNumber.PtgMul: ptg = new PtgMul(reader, ptgtype); break;
                            case PtgNumber.PtgDiv: ptg = new PtgDiv(reader, ptgtype); break;
                            case PtgNumber.PtgParen: ptg = new PtgParen(reader, ptgtype); break;
                            case PtgNumber.PtgNum: ptg = new PtgNum(reader, ptgtype); break;
                            case PtgNumber.PtgRef: ptg = new PtgRef(reader, ptgtype); break;
                            case PtgNumber.PtgRefN: ptg = new PtgRefN(reader, ptgtype); break;
                            case PtgNumber.PtgPower: ptg = new PtgPower(reader, ptgtype); break;
                            case PtgNumber.PtgPercent: ptg = new PtgPercent(reader, ptgtype); break;
                            case PtgNumber.PtgBool: ptg = new PtgBool(reader, ptgtype); break;
                            case PtgNumber.PtgGt: ptg = new PtgGt(reader, ptgtype); break;
                            case PtgNumber.PtgGe: ptg = new PtgGe(reader, ptgtype); break;
                            case PtgNumber.PtgLt: ptg = new PtgLt(reader, ptgtype); break;
                            case PtgNumber.PtgLe: ptg = new PtgLe(reader, ptgtype); break;
                            case PtgNumber.PtgEq: ptg = new PtgEq(reader, ptgtype); break;
                            case PtgNumber.PtgNe: ptg = new PtgNe(reader, ptgtype); break;
                            case PtgNumber.PtgUminus: ptg = new PtgUminus(reader, ptgtype); break;
                            case PtgNumber.PtgUplus: ptg = new PtgUplus(reader, ptgtype); break;
                            case PtgNumber.PtgStr: ptg = new PtgStr(reader, ptgtype); break;
                            case PtgNumber.PtgConcat: ptg = new PtgConcat(reader, ptgtype); break;
                            case PtgNumber.PtgUnion: ptg = new PtgUnion(reader, ptgtype); break;
                            case PtgNumber.PtgIsect: ptg = new PtgIsect(reader, ptgtype); break;
                            case PtgNumber.PtgMemErr: ptg = new PtgMemErr(reader, ptgtype); break;
                            case PtgNumber.PtgArea: ptg = new PtgArea(reader, ptgtype); break;
                            case PtgNumber.PtgAreaN: ptg = new PtgAreaN(reader, ptgtype); break;
                            case PtgNumber.PtgFuncVar: ptg = new PtgFuncVar(reader, ptgtype); break;
                            case PtgNumber.PtgFunc: ptg = new PtgFunc(reader, ptgtype); break;
                            case PtgNumber.PtgExp: ptg = new PtgExp(reader, ptgtype); break;
                            case PtgNumber.PtgRef3d: ptg = new PtgRef3d(reader, ptgtype); break;
                            case PtgNumber.PtgArea3d: ptg = new PtgArea3d(reader, ptgtype); break;
                            case PtgNumber.PtgNameX: ptg = new PtgNameX(reader, ptgtype); break;
                            case PtgNumber.PtgMissArg: ptg = new PtgMissArg(reader, ptgtype); break;


                            default: break;
                        }
                    }
                    i += ptg.getLength() - 1;
                    ptgStack.Push(ptg);
                }
            }
            catch (Exception)
            {
                throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
            }
            
            return ptgStack; 
        }

        public static String parseVirtualPath(String path)
        {
            char x01 = (char)0x01;
            char x02 = (char)0x02;
            char x03 = (char)0x03;
            char x05 = (char)0x05;
            char x06 = (char)0x06;
            char x07 = (char)0x07;
            char x08 = (char)0x08;


            path = path.Trim();
            if (path[0] == x01 && path[1] == x01)
            {
                path = path.Substring(2);
            }
            else if (path[0] == x01 && path[1] == x05)
            {
                path = path.Substring(2);
            }
            else if (path[0] == x01 && path[1] == x02)
            {
                path = path.Substring(2);
            }
            else if (path[0] == x01 && path[1] == x06)
            {
                path = path.Substring(2);
            }
            else if (path[0] == x01 && path[1] == x07)
            {
                path = path.Substring(2);
            }
            else if (path[0] == x01 && path[1] == x08)
            {
                path = path.Substring(2);
            }
            else if (path[0] == x01)
            {
                path = path.Substring(1);
            }


            /// Replace 0x03 with \
            path = path.Replace(x03, '\\');
            /// replace ' ' with %20
            path = path.Replace(" ", "%20");
            return path; 
        }

        /// <summary>
        /// This method reads x bytes from a IStreamReader to get a string from this
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="cch"></param>
        /// <param name="grbit"></param>
        /// <returns></returns>
        public static string getStringFromBiffRecord(IStreamReader reader,int cch, int grbit)
        {
            string value = ""; 
            if (grbit == 0)
            {
                for (int i = 0; i < cch; i++)
                {
                    value += (char)reader.ReadByte();
                }
            }
            else
            {
                for (int i = 0; i < cch; i++)
                {
                    value += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
                }
            }
            return value; 
        }
    }
}
