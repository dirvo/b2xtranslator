using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(0xF00A)]
    public class Shape : Record
    {
        public Int32 spid;

        /// <summary>
        /// This shape is a group shape 
        /// </summary>
        public bool fGroup;

        /// <summary>
        /// Not a top-level shape 
        /// </summary>
        public bool fChild;

        /// <summary>
        /// This is the topmost group shape.<br/>
        /// Exactly one of these per drawing. 
        /// </summary>
        public bool fPatriarch; 

        /// <summary>
        /// The shape has been deleted 
        /// </summary>
        public bool fDeleted;

        /// <summary>
        /// The shape is an OLE object 
        /// </summary>
        public bool fOleShape;

        /// <summary>
        /// Shape has a hspMaster property 
        /// </summary>
        public bool fHaveMaster;

        /// <summary>
        /// Shape is flipped horizontally 
        /// </summary>
        public bool fFlipH;

        /// <summary>
        /// Shape is flipped vertically 
        /// </summary>
        public bool fFlipV;

        /// <summary>
        /// Connector type of shape 
        /// </summary>
        public bool fConnector;

        /// <summary>
        /// Shape has an anchor of some kind 
        /// </summary>
        public bool fHaveAnchor;

        /// <summary>
        /// Background shape 
        /// </summary>
        public bool fBackground;

        /// <summary>
        /// Shape has a shape type property
        /// </summary>
        public bool fHaveSpt;

        /// <summary>
        /// The shape type of the shape
        /// </summary>
        public ShapeType ShapeType;

        public Shape(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.spid = this.Reader.ReadInt32();

            UInt32 flag = this.Reader.ReadUInt32();
            this.fGroup = Utils.BitmaskToBool(flag, 0x1);
            this.fChild = Utils.BitmaskToBool(flag, 0x2);
            this.fPatriarch = Utils.BitmaskToBool(flag, 0x4);
            this.fDeleted = Utils.BitmaskToBool(flag, 0x8);
            this.fOleShape = Utils.BitmaskToBool(flag, 0x10);
            this.fHaveMaster = Utils.BitmaskToBool(flag, 0x20);
            this.fFlipH = Utils.BitmaskToBool(flag, 0x40);
            this.fFlipV = Utils.BitmaskToBool(flag, 0x80);
            this.fConnector = Utils.BitmaskToBool(flag, 0x100);
            this.fHaveAnchor = Utils.BitmaskToBool(flag, 0x200);
            this.fBackground = Utils.BitmaskToBool(flag, 0x400);
            this.fHaveSpt = Utils.BitmaskToBool(flag, 0x800);

            this.ShapeType = ShapeType.GetShapeType(this.Instance);
        }

    }
}
