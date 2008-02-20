using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class ParagraphPropertyExceptions
    {
        /// <summary>
        /// Index to style descriptor of the style from which the 
        /// paragraph inherits its paragraph and character properties
        /// </summary>
        public UInt16 istd;

        /// <summary>
        /// A list of the sprms that encode the differences between 
        /// PAP for a paragraph and the PAP for the style used.
        /// </summary>
        public List<SinglePropertyModifier> grpprl;

        /// <summary>
        /// Creates a PAPX wich doesn't modify anything.<br/>
        /// The grpprl list is null
        /// </summary>
        public ParagraphPropertyExceptions()
        {
            grpprl = new List<SinglePropertyModifier>();
        }

        /// <summary>
        /// Parses the bytes to retrieve a PAPX
        /// </summary>
        /// <param name="bytes">The bytes starting with the istd</param>
        public ParagraphPropertyExceptions(byte[] bytes)
        {
            if (bytes.Length != 0)
            {
                this.istd = System.BitConverter.ToUInt16(bytes, 0);

                //read the sprms
                grpprl = new List<SinglePropertyModifier>();
                int sprmStart = 2;
                bool goOn = true;
                while (goOn)
                {
                    try
                    {
                        //make spra
                        UInt16 opCode = System.BitConverter.ToUInt16(bytes, sprmStart);
                        byte spra = (byte)((Int32)opCode >> 13);

                        // get size of operand
                        byte opSize = SinglePropertyModifier.GetOperandSize(spra);
                        byte lenByte = 0;
                        if (opSize == 255)
                        {
                            //the variable length stand in the byte after the opcode
                            lenByte = 1;
                            opSize = bytes[sprmStart + 2];
                        }

                        //copy sprm to array
                        byte[] sprm = new byte[2 + lenByte + opSize];
                        Array.Copy(bytes, sprmStart, sprm, 0, sprm.Length);

                        //parse and save
                        grpprl.Add(new SinglePropertyModifier(sprm));

                        sprmStart += sprm.Length;
                    }
                    catch (ArgumentException)
                    {
                        goOn = false;
                    }
                }
            }
            else
            {
                throw new ByteParseException("PAPX");
            }
        }
    }
}
