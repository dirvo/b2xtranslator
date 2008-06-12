using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class OfficeDrawingTable : Dictionary<Int32, FileShapeAddress>
    {
        public enum OfficeDrawingTableType
        {
            Header,
            MainDocument
        }

        private const int FSPA_LENGTH = 26;

        public OfficeDrawingTable(WordDocument doc, OfficeDrawingTableType type)
        {
            VirtualStreamReader reader = new VirtualStreamReader(doc.TableStream);

            //FSPA has size 26 + 4 byte for the FC = 30 byte
            int n = 0;
            int startFc = 0;
            if(type == OfficeDrawingTableType.MainDocument)
            {
                startFc = doc.FIB.fcPlcspaMom;
                n = (int)Math.Floor((double)doc.FIB.lcbPlcspaMom / 30);
            }
            else if(type == OfficeDrawingTableType.Header)
            {
                startFc = doc.FIB.fcPlcspaHdr;
                n = (int)Math.Floor((double)doc.FIB.lcbPlcspaHdr / 30);
            }
            
            //there are n+1 FCs ...
            doc.TableStream.Seek(startFc, System.IO.SeekOrigin.Begin);
            Int32[] fcs = new Int32[n+1];
            for (int i = 0; i < (n+1); i++)
            {
                fcs[i] = reader.ReadInt32();
            }

            //followed by n FSPAs
            for (int i = 0; i < n; i++)
            {
                FileShapeAddress fspa = null;
                if (type == OfficeDrawingTableType.Header)
                {
                    fspa = new FileShapeAddress(reader, doc.DrawingObjectTable);
                }
                else if (type == OfficeDrawingTableType.MainDocument)
                {
                    fspa = new FileShapeAddress(reader, doc.DrawingObjectTable);
                }
                this.Add(fcs[i], fspa);
            }
        }
    }
}
