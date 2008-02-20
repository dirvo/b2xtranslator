using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class SinglePropertyModifier
    {
        /// <summary>
        /// Identifies the type of a SPRM
        /// </summary>
        public enum SprmType
        {
            PAP = 1,
            CHP = 2,
            PIC = 3,
            SEP = 4,
            TAP = 5
        }

        /// <summary>
        /// The operation code identifies the property of the 
        /// PAP/CHP/PIC/SEP/TAP which sould be modified
        /// </summary>
        public UInt16 OpCode;

        /// <summary>
        /// This SPRM requires special handling
        /// </summary>
        public bool fSpec;

        /// <summary>
        /// The type of the SPRM
        /// </summary>
        public SprmType Type;

        /// <summary>
        /// The arguments which is applied to the property
        /// </summary>
        public byte[] Arguments;

        /// <summary>
        /// parses the byte to retrieve a SPRM
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public SinglePropertyModifier(byte[] bytes)
        {
            //first 2 bytes are the operation code ...
            this.OpCode = System.BitConverter.ToUInt16(bytes, 0);

            //... whereof bit 9 is fSpec ...
            UInt32 j = (UInt32)this.OpCode << 22;
            j = j >> 31;
            if (j == 1)
                this.fSpec = true;
            else
                this.fSpec = false;

            //... and bits 10,11,12 are the type ...
            UInt32 i = (UInt32)this.OpCode << 19;
            i = i >> 29;
            this.Type = (SprmType)i;

            //... and last 3 bits are the spra
            byte spra = (byte)((Int32)this.OpCode >> 13);
            byte opSize = GetOperandSize(spra);
            if (opSize == 255)
            {
                //the variable length stand in the byte after the opcode (byte2)
                opSize = bytes[2];
                //and the arguments start at the byte after that (byte3)
                this.Arguments = new byte[opSize];
                Array.Copy(bytes, 3, this.Arguments, 0, Arguments.Length);
            }
            else
            {
                this.Arguments = new byte[opSize];
                Array.Copy(bytes, 2, this.Arguments, 0, Arguments.Length);
            }
        }

        /// <summary>
        /// Get be used to get the size of the sprm's operand.
        /// Returns 0 if the Operation failed and 255 if the size is variable
        /// </summary>
        /// <param name="spra">the 3 bits for spra (as byte)</param>
        /// <returns>the size (as byte)</returns>
        public static byte GetOperandSize(byte spra)
        {
            switch (spra)
            {
                case 0: return 1;
                case 1: return 1;
                case 2: return 2;
                case 3: return 4;
                case 4: return 2;
                case 5: return 2;
                case 6: return 255;
                case 7: return 3;
                default: return 0;
            }
        }
    }
}
