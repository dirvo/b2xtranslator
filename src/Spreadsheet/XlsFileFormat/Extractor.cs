using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using System.Collections.ObjectModel;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;


namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{

    struct BiffHeader
    {
        public RecordNumber id;
        public UInt16 length;
    }

    public abstract class Extractor
    {
        public VirtualStreamReader StreamReader;   

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="sum">workbookstream </param>
        public Extractor(VirtualStreamReader reader)
        {
            this.StreamReader = reader;
            if (StreamReader == null)
            {
                throw new ExtractorException(ExtractorException.NULLPOINTEREXCEPTION);
            }
        }

        public abstract void extractData(); 
    }
}
