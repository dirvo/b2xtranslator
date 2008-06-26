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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// This class should extract the specific worksheet data. 
    /// </summary>
    public class BoundSheetExtractor : Extractor
    {
        /// <summary>
        /// Datacontainer for the worksheet
        /// </summary>
        private BoundSheetData bsd;

        /// <summary>
        /// CTor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="bsd"> Boundsheetdata container</param>
        public BoundSheetExtractor(VirtualStreamReader reader, BoundSheetData bsd)
            : base(reader) 
        {
            this.bsd = bsd;
            this.extractData(); 
        }

        /// <summary>
        /// Extracting the data from the stream 
        /// </summary>
        public override void extractData()
        {
            BiffHeader bh;
            try
            {
                while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
                {
                    bh.id = (RecordNumber)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();

                    // TraceLogger.Debug("BIFF {0}\t{1}\t", bh.id, bh.length);

                    if (bh.id == RecordNumber.EOF)
                    {
                        this.StreamReader.BaseStream.Seek(0, SeekOrigin.End); 
                    }
                    else if (bh.id == RecordNumber.LABELSST)
                    {
                        LABELSST labelsst = new LABELSST(this.StreamReader, bh.id, bh.length);
                        this.bsd.addLabelSST(labelsst); 
                    }
                    else if (bh.id == RecordNumber.MULRK) 
                    {
                        MULRK mulrk = new MULRK(this.StreamReader, bh.id, bh.length);
                        this.bsd.addMULRK(mulrk);
                    }
                    else if (bh.id == RecordNumber.NUMBER)
                    {
                        NUMBER number = new NUMBER(this.StreamReader, bh.id, bh.length);
                        this.bsd.addNUMBER(number);
                    }
                    else if (bh.id == RecordNumber.RK)
                    {
                        RK rk = new RK(this.StreamReader, bh.id, bh.length);
                        this.bsd.addRK(rk);
                    }
                    else if (bh.id == RecordNumber.MERGECELLS)
                    {
                        MERGECELLS mergecells = new MERGECELLS(this.StreamReader, bh.id, bh.length);
                        this.bsd.MERGECELLSData = mergecells;
                    }
                    else if (bh.id == RecordNumber.BLANK)
                    {
                        BLANK blankcell = new BLANK(this.StreamReader, bh.id, bh.length);
                        this.bsd.addBLANK(blankcell);
                    }
                    else if (bh.id == RecordNumber.MULBLANK)
                    {
                        MULBLANK mulblank = new MULBLANK(this.StreamReader, bh.id, bh.length);
                        this.bsd.addMULBLANK(mulblank);
                    }
                    else 
                    {   
                        // this else statement is used to read BiffRecords which aren't implemented 
                        byte[] buffer = new byte[bh.length];
                        buffer = this.StreamReader.ReadBytes(bh.length);               
                    }
                }
            }
            catch (Exception ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
        }
    }
}
