using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public abstract class PlexStruct
    {
        protected VirtualStreamReader _reader;

        public PlexStruct(VirtualStreamReader reader) 
        {
            this._reader = reader;
        }
    }
}
