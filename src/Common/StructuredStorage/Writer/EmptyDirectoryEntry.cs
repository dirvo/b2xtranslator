using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    /// <summary>
    /// Empty directory entry used to pad out directory stream.
    /// Author: math
    /// </summary>
    class EmptyDirectoryEntry : BaseDirectoryEntry
    {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="context">the current context</param>
        internal EmptyDirectoryEntry(StructuredStorageContext context)
            : base("", context)
        {
            Color = DirectoryEntryColor.DE_RED; // 0x0
            Type = DirectoryEntryType.STGTY_INVALID;            
        }

    }
}
