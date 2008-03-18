using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TableInfo
    {
        /// <summary>
        /// 
        /// </summary>
        public bool fInTable;

        /// <summary>
        /// 
        /// </summary>
        public bool fTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTableCell;

        /// <summary>
        /// 
        /// </summary>
        public UInt32 iTap;

        public TableInfo(ParagraphPropertyExceptions papx)
        {
            foreach (SinglePropertyModifier sprm in papx.grpprl)
            {
                if (sprm.OpCode == 0x2416)
                {
                    this.fInTable = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == 0x2417)
                {
                    this.fTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == 0x244B)
                {
                    this.fInnerTableCell = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == 0x244C)
                {
                    this.fInnerTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == 0x6649)
                {
                    this.iTap = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (this.iTap > 0)
                        this.fInTable = true;
                }
                if (sprm.OpCode == 0x66A)
                {
                    //add value!
                    this.iTap = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (this.iTap > 0)
                        this.fInTable = true;
                }
            }
        }
    }
}
