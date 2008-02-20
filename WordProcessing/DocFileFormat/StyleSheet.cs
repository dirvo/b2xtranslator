using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.WordFileFormat
{
    public class StyleSheet
    {
        /// <summary>
        /// The StyleSheetInformation
        /// </summary>
        public StyleSheetInformation stshi;

        /// <summary>
        /// The list contains all styles
        /// </summary>
        public List<StyleSheetDescription> Styles;

        /// <summary>
        /// Parses the streams to retrieve a StyleSheet
        /// </summary>
        /// <param name="fib">The FileInformationBlock</param>
        /// <param name="tableStream">The 0Table or 1Table stream</param>
        public StyleSheet(FileInformationBlock fib, VirtualStream tableStream)
        {
            //read size of the STSHI
            byte[] stshiLengthBytes = new byte[2];
            tableStream.Read(stshiLengthBytes, stshiLengthBytes.Length, fib.fcStshf);
            Int16 cbStshi = System.BitConverter.ToInt16(stshiLengthBytes, 0);

            //read the bytes of the stshi
            byte[] stshi = new byte[cbStshi];
            tableStream.Read(stshi, cbStshi, fib.fcStshf + 2);

            //parses STSHI
            this.stshi = new StyleSheetInformation(stshi);

            //create list for STDs
            this.Styles = new List<StyleSheetDescription>();
            for (int i = 0; i < this.stshi.cstd; i++)
            {
                //get the cbStd
                byte[] cbStdBytes = new byte[2];
                tableStream.Read(cbStdBytes);
                UInt16 cbStd = System.BitConverter.ToUInt16(cbStdBytes, 0);

                if (cbStd != 0)
                {
                    //read the STD bytes
                    byte[] std = new byte[cbStd];
                    tableStream.Read(std);

                    //parse the STD
                    this.Styles.Add(new StyleSheetDescription(std));
                }
                else
                {
                    this.Styles.Add(null);
                }
            }

        }
    }
}
