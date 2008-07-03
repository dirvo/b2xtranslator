using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FileShapeAddress
    {
        public enum AnchorType
        {
            margin,
            page,
            text
        }

        /// <summary>
        /// Shape Identifier. Used in conjunction with the office art data 
        /// (found via fcDggInfo in the FIB) to find the actual data for this shape.
        /// </summary>
        public Int32 spid;

        /// <summary>
        /// Left of rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public Int32 xaLeft;

        /// <summary>
        /// Top of rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public Int32 yaTop;

        /// <summary>
        /// Right of rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public Int32 xaRight;

        /// <summary>
        /// Bottom of the rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public Int32 yaBottom;

        /// <summary>
        /// true in the undo doc when shape is from the header doc<br/>
        /// false otherwise (undefined when not in the undo doc)
        /// </summary>
        public bool fHdr;

        /// <summary>
        /// X position of shape relative to anchor CP<br/>
        /// 0 relative to page margin<br/>
        /// 1 relative to top of page<br/>
        /// 2 relative to text (column for horizontal text; paragraph for vertical text)<br/>
        /// 3 reserved for future use
        /// </summary>
        public AnchorType bx;

        /// <summary>
        /// Y position of shape relative to anchor CP<br/>
        /// 0 relative to page margin<br/>
        /// 1 relative to top of page<br/>
        /// 2 relative to text (column for horizontal text; paragraph for vertical text)<br/>
        /// 3 reserved for future use
        /// </summary>
        public AnchorType by;

        /// <summary>
        /// Text wrapping mode <br/>
        /// 0 like 2, but doesn�t require absolute object <br/>
        /// 1 no text next to shape <br/>
        /// 2 wrap around absolute object <br/>
        /// 3 wrap as if no object present <br/>
        /// 4 wrap tightly around object <br/>
        /// 5 wrap tightly, but allow holes <br/>
        /// 6-15 reserved for future use
        /// </summary>
        public UInt16 wr;

        /// <summary>
        /// Text wrapping mode type (valid only for wrapping modes 2 and 4)<br/>
        /// 0 wrap both sides <br/>
        /// 1 wrap only on left <br/>
        /// 2 wrap only on right <br/>
        /// 3 wrap only on largest side
        /// </summary>
        public UInt16 wrk;

        /// <summary>
        /// When set, temporarily overrides bx, by, 
        /// forcing the xaLeft, xaRight, yaTop, and yaBottom fields 
        /// to all be page relative.
        /// </summary>
        public bool fRcaSimple;

        /// <summary>
        /// true: shape is below text <br/>
        /// false: shape is above text
        /// </summary>
        public bool fBelowText;

        /// <summary>
        /// true: anchor is locked <br/>
        /// fasle: anchor is not locked
        /// </summary>
        public bool fAnchorLock;

        /// <summary>
        /// Count of textboxes in shape (undo doc only)
        /// </summary>
        public Int32 cTxbx;

        /// <summary>
        /// The shapecontainer
        /// </summary>
        public ShapeContainer ShapeContainer;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reader"></param>
        public FileShapeAddress(VirtualStreamReader reader, OfficeArtContent drawingTable)
        {
            this.spid = reader.ReadInt32();
            this.xaLeft = reader.ReadInt32();
            this.yaTop = reader.ReadInt32();
            this.xaRight = reader.ReadInt32();
            this.yaBottom = reader.ReadInt32();

            UInt16 flag = reader.ReadUInt16();
            this.fHdr = Tools.Utils.BitmaskToBool(flag, 0x0001);
            this.bx = (AnchorType)Tools.Utils.BitmaskToInt(flag, 0x0006);
            this.by = (AnchorType)Tools.Utils.BitmaskToInt(flag, 0x0018);
            this.wr = (UInt16)Tools.Utils.BitmaskToInt(flag, 0x01E0);
            this.wrk = (UInt16)Tools.Utils.BitmaskToInt(flag, 0x1E00);
            this.fRcaSimple = Tools.Utils.BitmaskToBool(flag, 0x2000);
            this.fBelowText = Tools.Utils.BitmaskToBool(flag, 0x4000);
            this.fAnchorLock = Tools.Utils.BitmaskToBool(flag, 0x8000);

            this.cTxbx = reader.ReadInt32();

            this.ShapeContainer = drawingTable.GetShapeContainer(this.spid);
        }
    }
}
