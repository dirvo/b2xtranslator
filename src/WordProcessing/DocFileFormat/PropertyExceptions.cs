using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public abstract class PropertyExceptions : IVisitable
    {
        /// <summary>
        /// A list of the sprms that encode the differences between 
        /// CHP for a character and the PAP for the paragraph style used.
        /// </summary>
        public List<SinglePropertyModifier> grpprl;

        public PropertyExceptions()
        {
            this.grpprl = new List<SinglePropertyModifier>();
        }

        public PropertyExceptions(byte[] bytes)
        {
            this.grpprl = new List<SinglePropertyModifier>();

            if (bytes.Length != 0)
            {
                //read the sprms
                
                int sprmStart = 0;
                bool goOn = true;
                while (goOn)
                {
                    //enough bytes to read?
                    if (sprmStart + 2 < bytes.Length)
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

                        if (bytes.Length >= sprmStart + sprm.Length)
                        {
                            Array.Copy(bytes, sprmStart, sprm, 0, sprm.Length);

                            //parse and save
                            grpprl.Add(new SinglePropertyModifier(sprm));

                            sprmStart += sprm.Length;
                        }
                        else
                        {
                            goOn = false;
                        }
                    }
                    else
                    {
                        goOn = false;
                    }
                }
            }
        }

        public abstract void Convert<T>(T mapping);
    }

}
