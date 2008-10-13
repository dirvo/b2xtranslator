using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class CommandTable
    {
        public StringTable CommandStringTable;

        public List<MacroData> MacroDatas;

        public Dictionary<Int32, String> MacroNames;

        bool breakWhile;

        public CommandTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            tableStream.Seek(fib.fcCmds, System.IO.SeekOrigin.Begin);
            VirtualStreamReader reader = new VirtualStreamReader(tableStream);

            //byte[] bytes = reader.ReadBytes((int)fib.lcbCmds);

            //skip the version
            reader.ReadByte();

            //parse the commandtable
            while (reader.BaseStream.Position < (fib.fcCmds + fib.lcbCmds) && !breakWhile)
            {
                //read the type
                byte ch = reader.ReadByte();

                switch (ch)
                {
                    case 0x1:
                        //it's a PlfMcd
                        this.MacroDatas = new List<MacroData>();
                        int iMacMcd = reader.ReadInt32();
                        for (int i = 0; i < iMacMcd; i++)
                        {
                            this.MacroDatas.Add(new MacroData(reader, 24));
                        }
                        break;
                    case 0x2:
                        //it's a PlfAcd

                        //skip the ACDs
                        int iMacAcd = reader.ReadInt32();
                        reader.ReadBytes(iMacAcd * 4);
                        break;
                    case 0x3:
                    case 0x4:
                        //it's a PlfKme

                        //skip the KMEs
                        int iMacKme = reader.ReadInt32();
                        reader.ReadBytes(iMacKme * 14);
                        break;
                    case 0x10:
                        //it's a TcgSttbf
                        this.CommandStringTable = new StringTable(typeof(String), reader); 
                        break;
                    case 0x11:
                        //it's a MacroNames table
                        this.MacroNames = new Dictionary<int, string>();
                        int iMacMn = reader.ReadInt16();
                        for (int i = 0; i < iMacMn; i++)
                        {
                            Int16 ibst = reader.ReadInt16();
                            Int16 cch = reader.ReadInt16();
                            this.MacroNames[ibst] = Encoding.Unicode.GetString(reader.ReadBytes(cch * 2));
                            //skip the terminating zero
                            reader.ReadBytes(2);
                        }
                        break;
                    case 0x12:
                        //it's a CTBWRAPPER structure
                        reader.ReadBytes(8);
                        Int16 cbTBD = reader.ReadInt16();
                        Int16 cCust = reader.ReadInt16();
                        Int32 cbDTBC = reader.ReadInt32();
                        //skip the rtbdc
                        reader.ReadBytes(cbDTBC);
                        break;
                    default:
                        breakWhile = true;
                        break;
                }
            }
        }
    }
}
