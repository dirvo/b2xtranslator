using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using System.Collections.ObjectModel;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.XlsFileFormat.Exceptions;
using XlsFileFormat; 

namespace DIaLOGIKa.b2xtranslator.XlsFileFormat
{
    public class FileInformationExtractor
    {
        public VirtualStream SummaryStream;         // Summary stream 

        public string Title;

        public string buffer; 

        struct BiffHeader
        {
            public RecordNumber id;
            public UInt16 length;
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="sum">Summary stream </param>
        public FileInformationExtractor(VirtualStream sum)
        {
            this.Title = null; 
            if (sum == null)
            {
                throw new ExtractorException(ExtractorException.NULLPOINTEREXCEPTION); 
            }
            this.SummaryStream = sum;
            this.extractData(); 


        }

        /// <summary>
        /// Extracts the data from the stream 
        /// </summary>
        public void extractData()
        {
            BiffHeader bh;
            StreamWriter sw = null;
            sw = new StreamWriter(Console.OpenStandardOutput());
            try
            {
                while ((ulong)this.SummaryStream.Position < this.SummaryStream.SizeOfStream)
                {
                    bh.id = (RecordNumber)this.SummaryStream.ReadUInt16();
                    bh.length = this.SummaryStream.ReadUInt16();

                    byte[] buf = new byte[bh.length];
                    if (bh.length != this.SummaryStream.Read(buf, bh.length))
                        sw.WriteLine("EOF");

                    sw.Write("BIFF {0}\t{1}\t", bh.id, bh.length);
                    //Dump(buffer);
                    int count = 0;
                    foreach (byte b in buf)
                    {
                        sw.Write("{0:X02} ", b);
                        count++;
                        if (count % 16 == 0 && count < buf.Length)
                            sw.Write("\n\t\t\t");
                    }
                    sw.Write("\n");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            this.buffer = sw.ToString();
         }

        public override string ToString()
        {
            string returnvalue = "Title: " + this.Title;
            return returnvalue; 
        }
    }
}
