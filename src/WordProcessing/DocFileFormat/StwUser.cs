using System;
using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    /// <summary>
    /// The StwUser structure specifies the names and values of the user-defined 
    /// variables that are stored in the document
    /// </summary>
    public class StwUser : Dictionary<String, String>
    {
        /// <summary>
        /// Parses an StwUser structure into a dictionary of key-value-pairs
        /// </summary>
        /// <param name="tableStream">The input table stream</param>
        /// <param name="fcStwUser">fcStwUser (4 bytes): An unsigned integer that specifies an offset into the Table Stream. 
        /// An StwUser that specifies the user-defined variables and VBA digital signature (2), as specified by 
        /// [MS-OSHARED] section 2.3.2, begins at this offset. 
        /// 
        /// If lcbStwUser is zero, fcStwUser is undefined and MUST be ignored.</param>
        /// <param name="lcbStwUser">lcbStwUser (4 bytes): An unsigned integer that specifies the size, in bytes, 
        /// of the StwUser at offset fcStwUser</param>
        public StwUser(VirtualStream tableStream, UInt32 fcStwUser, UInt32 lcbStwUser)
        {
            if (lcbStwUser == 0)
            {
                return;
            }
            tableStream.Seek(fcStwUser, System.IO.SeekOrigin.Begin);

            // parse the names
            var names = new StringTable(typeof(string), tableStream, fcStwUser, lcbStwUser);

            // parse the values
            var values = new List<string>();
            while (tableStream.Position < fcStwUser+lcbStwUser)
            {
                values.Add(Utils.ReadXst(tableStream));
            }

            // map to the dictionary
            if (names.Strings.Count == values.Count)
            {
                for (int i = 0; i < names.Strings.Count; i++)
                {
                    this.Add(names.Strings[i], values[i]);
                }
            }
        }
    }
}
