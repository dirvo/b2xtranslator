using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class MacroData : ByteStructure
    {
        /// <summary>
        /// Unsigned integer that specifies the name of the macro.<br/>
        /// Macro name is specified by MacroName.xstz of the MacroName entry in 
        /// the MacroNames such that MacroName.ibst equals ibst. <br/>
        /// MacroNames MUST contain such an entry.
        /// </summary>
        public Int16 ibst;

        /// <summary>
        /// An unsigned integer that specifies the index into the 
        /// CommandStringTable (CommandTable.CommandStringTable)
        /// where the macro‘s name and arguments are specified.
        /// </summary>
        public Int16 ibstName;

        public MacroData(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            //first 2 bytes are reserved
            reader.ReadBytes(2);

            this.ibst = reader.ReadInt16();

            this.ibstName = reader.ReadInt16();

            //last 18 bytes are reserved
            reader.ReadBytes(18);
        }
    }
}
