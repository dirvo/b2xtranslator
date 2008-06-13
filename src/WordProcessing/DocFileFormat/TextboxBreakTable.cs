using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class TextboxBreakTable
    {
        public enum TextboxBreakTableType
        {
            Header,
            MainDocument
        }

        public List<Int32> CharacterPositions;
        public List<BreakDescriptor> Breaks;

        private const int BKD_LENGTH = 6;

        public TextboxBreakTable(FileInformationBlock fib, VirtualStream stream, TextboxBreakTableType type)
        {
            BinaryReader reader = new BinaryReader(stream);
            int n;

            if (type == TextboxBreakTableType.MainDocument)
            {
                stream.Seek(fib.fcPlcftxbxBkd, System.IO.SeekOrigin.Begin);
                n = (int)Math.Floor((double)fib.lcbPlcftxbxBkd / (BKD_LENGTH + 4));
            }
            else 
            {
                stream.Seek(fib.fcPlcftxbxHdrBkd, System.IO.SeekOrigin.Begin);
                n = (int)Math.Floor((double)fib.lcbPlcftxbxHdrBkd / (BKD_LENGTH + 4));
            }

            
            //there are n+1 FCs ...
            this.CharacterPositions = new List<Int32>();
            for (int i = 0; i < (n + 1); i++)
            {
                this.CharacterPositions.Add(reader.ReadInt32());
            }

            //followed by n BKDs
            this.Breaks = new List<BreakDescriptor>();
            for (int i = 0; i < n; i++)
            {
                BreakDescriptor bkd = new BreakDescriptor(reader);
                this.Breaks.Add(bkd);
            }
        }
    }
}
