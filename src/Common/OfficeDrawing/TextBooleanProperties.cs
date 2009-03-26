using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class TextBooleanProperties
    {
        public bool fFitShapeToText;
        public bool fAutoTextMargin;
        public bool fSelectText;
        public bool fUsefFitShapeToText;
        public bool fUsefAutoTextMargin;
        public bool fUsefSelectText;

        public TextBooleanProperties(UInt32 entryOperand)
        {
            //1 is unused
            fFitShapeToText = Utils.BitmaskToBool(entryOperand, 0x100000 >> 0);
            //1 is unused
            fAutoTextMargin = Utils.BitmaskToBool(entryOperand, 0x100000 >> 2);
            fSelectText = Utils.BitmaskToBool(entryOperand, 0x100000 >> 3);
            //12 unused
            fUsefFitShapeToText = Utils.BitmaskToBool(entryOperand, 0x100000 >> 16);
            //1 is unused
            fUsefAutoTextMargin = Utils.BitmaskToBool(entryOperand, 0x100000 >> 18);
            fUsefSelectText = Utils.BitmaskToBool(entryOperand, 0x100000 >> 19);
        }
    }
}
