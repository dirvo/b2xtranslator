/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;

namespace DIaLOGIKa.b2xtranslator.Tools
{
    public class Utils
    {
        public static int ArraySum(byte[] values)
        {
            int ret = 0;
            foreach (byte b in values)
            {
                ret += b;
            }
            return ret;
        }

        public static bool BitmaskToBool(Int32 value, Int32 mask)
        {
            return ((value & mask) == mask);
        }

        public static bool BitmaskToBool(UInt32 value, UInt32 mask)
        {
            return ((value & mask) == mask);
        }

        public static byte BitmaskToByte(Int32 value, Int32 mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
                value = value >> 1;
            return Convert.ToByte(value);
        }

        public static Int32 BitmaskToInt32(Int32 value, Int32 mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
                value = value >> 1;
            return value;
        }

        public static UInt32 BitmaskToUInt32(UInt32 value, UInt32 mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
                value = value >> 1;
            return value;
        }

        public static bool IntToBool(int value)
        {
            if (value == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        public static bool ByteToBool(byte value)
        {
            if (value == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static char[] ClearCharArray(char[] values)
        {
            char[] ret = new char[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = Convert.ToChar(0);
            }
            return ret;
        }

        public static int[] ClearIntArray(int[] values)
        {
            int[] ret = new int[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = 0;
            }
            return ret;
        }

        public static short[] ClearShortArray(ushort[] values)
        {
            short[] ret = new short[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = 0;
            }
            return ret;
        }

        public static UInt32 BitArrayToUInt32(BitArray bits)
        {
            double ret = 0;
            for (int i = 0; i < bits.Count; i++)
            {
                if (bits[i])
                {
                    ret += Math.Pow((double)2, (double)i);
                }
            }
            return (UInt32)ret;
        }

        public static BitArray BitArrayCopy(BitArray source, int sourceIndex, int copyCount)
        {
            bool[] ret = new bool[copyCount];

            int j = 0;
            for (int i = sourceIndex; i < (copyCount + sourceIndex); i++)
            {
                ret[j] = source[i];
                j++;
            }

            return new BitArray(ret);
        }

        public static string GetHashDump(byte[] bytes)
        {
            int colCount = 16;
            string ret = String.Format("({0:X04}) ", 0);

            int colCounter = 0;
            for (int i = 0; i < bytes.Length; i++)
            {
                if (colCounter == colCount)
                {
                    colCounter = 0;
                    ret += "\n" + String.Format("({0:X04}) ", i);
                }
                ret += String.Format("{0:X02} ", bytes[i]);
                colCounter++;
            }

            return ret;
        }
    }
}
