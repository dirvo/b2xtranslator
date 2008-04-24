using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// Used for mapping Office record TypeCodes to the classes implementing them.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class OfficeRecordAttribute : Attribute
    {
        public UInt16 TypeCode;
    }
}
