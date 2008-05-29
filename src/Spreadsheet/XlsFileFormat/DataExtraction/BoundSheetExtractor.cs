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
            StreamWriter sw = null;

            sw = new StreamWriter(Console.OpenStandardOutput());
            try
            {
                while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
                {
                    bh.id = (RecordNumber)this.StreamReader.ReadUInt16();

                    bh.length = this.StreamReader.ReadUInt16();
                    if (bh.id == RecordNumber.EOF)
                    {
                        this.StreamReader.BaseStream.Seek(0, SeekOrigin.End); 
                        sw.Write("EOF");
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
                    else 
                    {   
                        byte[] buffer = new byte[bh.length];
                        buffer = this.StreamReader.ReadBytes(bh.length);
                        if (bh.length != buffer.Length)
                            sw.WriteLine("EOF");

                        sw.Write("BIFF {0}\t{1}\t", bh.id, bh.length);
                        //Dump(buffer);
                        int count = 0;
                        foreach (byte b in buffer)
                        {
                            sw.Write("{0:X02} ", b);
                            count++;
                            if (count % 16 == 0 && count < buffer.Length)
                                sw.Write("\n\t\t\t");
                        }
                        sw.Write("\n"); 
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            sw.Close();
        }
    }
}
