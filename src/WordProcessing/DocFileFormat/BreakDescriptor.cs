using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class BreakDescriptor : PlexStruct
    {
        /// <summary>
        /// Except in textbox BKD, index to PGD in plfpgd that describes the page this break is on
        /// </summary>
        public Int16 ipgd;

        /// <summary>
        /// In textbox BKD
        /// </summary>
        public Int16 itxbxs;

        /// <summary>
        /// Number of cp's considered for this break; note that the CP's described by cpDepend in this break reside in the next BKD
        /// </summary>
        public Int16 dcpDepend;

        /// <summary>
        /// 
        /// </summary>
        public UInt16 icol;

        /// <summary>
        /// When true, this indicates that this is a table break.
        /// </summary>
        public bool fTableBreak;

        /// <summary>
        /// When true, this indicates that this is a column break.
        /// </summary>
        public bool fColumnBreak;

        /// <summary>
        /// Used temporarily while Word is running.
        /// </summary>
        public bool fMarked;

        /// <summary>
        /// In textbox BKD, when true indicates cpLim of this textbox is not valid
        /// </summary>
        public bool fUnk;

        /// <summary>
        /// In textbox BKD, when true indicates that text overflows the end of this textbox
        /// </summary>
        public bool fTextOverflow;

        public BreakDescriptor(VirtualStreamReader reader) : base(reader)
        {
            this.ipgd = reader.ReadInt16();
            this.itxbxs = this.ipgd;
            this.dcpDepend = reader.ReadInt16();
            int flag = (int)reader.ReadInt16();
            this.icol = (UInt16)Utils.BitmaskToInt(flag, 0x00FF);
            this.fTableBreak = Utils.BitmaskToBool(flag, 0x0100);
            this.fColumnBreak = Utils.BitmaskToBool(flag, 0x0200);
            this.fMarked = Utils.BitmaskToBool(flag, 0x0400);
            this.fUnk = Utils.BitmaskToBool(flag, 0x0800);
            this.fTextOverflow = Utils.BitmaskToBool(flag, 0x1000);
        }
    }
}
