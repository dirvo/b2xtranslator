﻿using System;
using System.Collections.Generic;
using System.IO;
namespace DIaLOGIKa.b2xtranslator.StructuredStorageReader
{
    public interface IStructuredStorageFile : IDisposable
    {
        /// <summary>
        /// Collection of all entries contained in a compound file
        /// </summary> 
        ICollection<DirectoryEntry> AllEntries { get; }

        /// <summary> 
        /// Collection of all stream entries contained in a compound file
        /// </summary> 
        ICollection<DirectoryEntry> AllStreamEntries { get; }
        
        /// <summary>
        /// Collection of all entry names contained in a compound file
        /// </summary>        
        ICollection<string> FullNameOfAllEntries { get; }

        /// <summary>
        /// Collection of all stream entry names contained in a compound file
        /// </summary>        
        ICollection<string> FullNameOfAllStreamEntries { get; }

        /// <summary>
        /// Closes the file handle
        /// </summary>
        void Close();

        /// <summary>
        /// Returns a handle to a stream with the given name/path.
        /// If a path is used, it must be preceeded by '\'.
        /// The characters '\' ( if not separators in the path) and '%' must be masked by '%XXXX'
        /// where 'XXXX' is the unicode in hex of '\' and '%', respectively
        /// </summary>
        /// <param name="path">The path of the virtual stream.</param>
        /// <returns>An object which enables access to the virtual stream.</returns>
        VirtualStream GetStream(string path);
        // TODO: return a System.IO.Stream object only
    }
}
