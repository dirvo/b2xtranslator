using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class PictureDescriptor : IVisitable
    {
        public enum PictureType
        {
            jpg,
            png,
            wmf
        }

        public struct MetafilePicture
        {
            /// <summary>
            /// Specifies the mapping mode in which the picture is drawn.
            /// </summary>
            public Int16 mm;

            /// <summary>
            /// Specifies the size of the metafile picture for all modes except the MM_ISOTROPIC and MM_ANISOTROPIC modes.<br/>
            /// (For more information about these modes, see the yExt member.) <br/>
            /// The x-extent specifies the width of the rectangle within which the picture is drawn.<br/>
            /// The coordinates are in units that correspond to the mapping mode.<br/>
            /// </summary>
            public Int16 xExt;

            /// <summary>
            /// Specifies the size of the metafile picture for all modes except the MM_ISOTROPIC and MM_ANISOTROPIC modes.<br/> 
            /// The y-extent specifies the height of the rectangle within which the picture is drawn.<br/>
            /// The coordinates are in units that correspond to the mapping mode. <br/>
            /// For MM_ISOTROPIC and MM_ANISOTROPIC modes, which can be scaled, the xExt and yExt members 
            /// contain an optional suggested size in MM_HIMETRIC units.<br/>
            /// For MM_ANISOTROPIC pictures, xExt and yExt can be zero when no suggested size is supplied.<br/>
            /// For MM_ISOTROPIC pictures, an aspect ratio must be supplied even when no suggested size is given.<br/>
            /// (If a suggested size is given, the aspect ratio is implied by the size.)<br/>
            /// To give an aspect ratio without implying a suggested size, set xExt and yExt to negative values 
            /// whose ratio is the appropriate aspect ratio.<br/>
            /// The magnitude of the negative xExt and yExt values is ignored; only the ratio is used.
            /// </summary>
            public Int16 yExt;

            /// <summary>
            /// Handle to a memory metafile.
            /// </summary>
            public Int16 hMf;
        }

        /// <summary>
        /// Rectangle for window origin and extents when metafile is stored (ignored if 0).
        /// </summary>
        public byte[] rcWinMf;

        /// <summary>
        /// Horizontal measurement in twips of the rectangle the picture should be imaged within.
        /// </summary>
        public Int16 dxaGoal;

        /// <summary>
        /// Vertical measurement in twips of the rectangle the picture should be imaged within.
        /// </summary>
        public Int16 dyaGoal;

        /// <summary>
        /// Horizontal scaling factor supplied by user expressed in .001% units
        /// </summary>
        public UInt16 mx;

        /// <summary>
        /// Vertical scaling factor supplied by user expressed in .001% units
        /// </summary>
        public UInt16 my;

        /// <summary>
        /// The type of the picture
        /// </summary>
        public PictureType Type;

        /// <summary>
        /// The name of the picture
        /// </summary>
        public string Name;

        /// <summary>
        /// The data of the windows metafile picture (WMF)
        /// </summary>
        public MetafilePicture mfp;

        /// <summary>
        /// The amount the picture has been cropped on the left in twips
        /// </summary>
        public Int16 dxaCropLeft;

        /// <summary>
        /// The amount the picture has been cropped on the top in twips
        /// </summary>
        public Int16 dyaCropTop;

        /// <summary>
        /// The amount the picture has been cropped on the right in twips
        /// </summary>
        public Int16 dxaCropRight;

        /// <summary>
        /// The amount the picture has been cropped on the bottom in twips
        /// </summary>
        public Int16 dyaCropBottom;

        /// <summary>
        /// Border above picture
        /// </summary>
        public BorderCode brcTop;

        /// <summary>
        /// Border to the left of the picture
        /// </summary>
        public BorderCode brcLeft;

        /// <summary>
        /// Border below picture
        /// </summary>
        public BorderCode brcBottom;

        /// <summary>
        /// Border to the right of the picture
        /// </summary>
        public BorderCode brcRight;

        /// <summary>
        /// Horizontal offset of hand annotation origin
        /// </summary>
        public Int16 dxaOrigin;

        /// <summary>
        /// vertical offset of hand annotation origin
        /// </summary>
        public Int16 dyaOrigin;

        /// <summary>
        /// unused
        /// </summary>
        public Int16 cProps;

        public ShapeContainer ShapeContainer;

        public BlipStoreEntry BlipStoreEntry;

        /// <summary>
        /// Parses the CHPX for a fcPic an loads the PictureDescriptor at this offset
        /// </summary>
        /// <param name="chpx">The CHPX that holds a SPRM for fcPic</param>
        public PictureDescriptor(CharacterPropertyExceptions chpx, VirtualStream stream)
        {
            //Get start and length of the PICT
            Int32 fc = getFcPic(chpx);
            if (fc >= 0)
            {
                parse(stream, fc);
            }
        }

        private void parse(VirtualStream stream, Int32 fc)
        {
            stream.Seek(fc, System.IO.SeekOrigin.Begin);
            VirtualStreamReader reader = new VirtualStreamReader(stream);

            Int32 lcb = reader.ReadInt32();

            if (lcb > 0)
            {
                UInt16 cbHeader = reader.ReadUInt16();

                this.mfp = new MetafilePicture();
                this.mfp.mm = reader.ReadInt16();
                this.mfp.xExt = reader.ReadInt16();
                this.mfp.yExt = reader.ReadInt16();
                this.mfp.hMf = reader.ReadInt16();

                if (this.mfp.mm > 98)
                {
                    this.rcWinMf = reader.ReadBytes(14);

                    //dimensions
                    this.dxaGoal = reader.ReadInt16();
                    this.dyaGoal = reader.ReadInt16();
                    this.mx = reader.ReadUInt16();
                    this.my = reader.ReadUInt16();

                    //cropping
                    this.dxaCropLeft = reader.ReadInt16();
                    this.dyaCropTop = reader.ReadInt16();
                    this.dxaCropRight = reader.ReadInt16();
                    this.dyaCropBottom = reader.ReadInt16();

                    Int16 brcl = reader.ReadInt16();

                    //borders
                    this.brcTop = new BorderCode(reader.ReadBytes(4));
                    this.brcLeft = new BorderCode(reader.ReadBytes(4));
                    this.brcBottom = new BorderCode(reader.ReadBytes(4));
                    this.brcRight = new BorderCode(reader.ReadBytes(4));

                    this.dxaOrigin = reader.ReadInt16();
                    this.dyaOrigin = reader.ReadInt16();
                    this.cProps = reader.ReadInt16();

                    //Parse the OfficeDrawing Stuff
                    this.ShapeContainer = (ShapeContainer)Record.readRecord(reader);
                    this.BlipStoreEntry = (BlipStoreEntry)Record.readRecord(reader);
                }
            }
        }



        /// <summary>
        /// Returns the fcPic into the "data" stream, where the PIC begins.
        /// Returns -1 if the CHPX has no fcPic.
        /// </summary>
        /// <param name="chpx">The CHPX</param>
        /// <returns></returns>
        public static Int32 getFcPic(CharacterPropertyExceptions chpx)
        {
            Int32 ret = -1;
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                switch (sprm.OpCode)
	            {
                    case SinglePropertyModifier.OperationCode.sprmCPicLocation:
                        ret = System.BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmCHsp:
                        ret = System.BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmCFData:
                        break;
	            }
                
            }
            return ret;
        }


        #region IVisitable Members

        public virtual void Convert<T>(T mapping)
        {
            ((IMapping<PictureDescriptor>)mapping).Apply(this);
        }

        #endregion
    }
}
