using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    class EmptyDirectoryEntry : BaseDirectoryEntry
    {
        internal EmptyDirectoryEntry(StructuredStorageContext context)
            : base("", context)
        {
            Color = DirectoryEntryColor.DE_RED; // 0x0
            Type = DirectoryEntryType.STGTY_INVALID;            
        }

    }
}
