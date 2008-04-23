using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

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
        /// The bytes of the picture
        /// </summary>
        public byte[] Picture;

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

        /// <summary>
        /// Parses the CHPX for a fcPic an loads the PictureDescriptor at this offset
        /// </summary>
        /// <param name="chpx">The CHPX that holds a SPRM for fcPic</param>
        public PictureDescriptor(CharacterPropertyExceptions chpx, VirtualStream dataStream)
        {
            VirtualStreamReader dataReader = new VirtualStreamReader(dataStream);

            //Get start and length of the PICT
            Int32 fc = getFcPic(chpx);
            if (fc >= 0)
            {
                dataStream.Seek(fc, System.IO.SeekOrigin.Begin);
                Int32 lcb = dataReader.ReadInt32();

                //byte[] lcbBytes = new byte[4];
                //dataStream.Read(lcbBytes, 0, 4, fc);
                //Int32 lcb = System.BitConverter.ToInt32(lcbBytes, 0);

                if (lcb > 0)
                {
                    //read the bytes of the PIC
                    byte[] pictBytes = dataReader.ReadBytes(lcb);

                    //parse
                    parseBytes(pictBytes);
                }
            }
        }

        /// <summary>
        /// Parses the bytes to retrieve a PictureDescriptor
        /// </summary>
        /// <param name="bytes">The bytes beginng at cbHeader, including the picture</param>
        public PictureDescriptor(byte[] bytes)
        {
            parseBytes(bytes);
        }

        private void parseBytes(byte[] bytes)
        {
            UInt16 cbHeader = System.BitConverter.ToUInt16(bytes, 0);

            this.mfp = new MetafilePicture();
            this.mfp.mm = System.BitConverter.ToInt16(bytes, 2);
            this.mfp.xExt = System.BitConverter.ToInt16(bytes, 4);
            this.mfp.yExt = System.BitConverter.ToInt16(bytes, 6);
            this.mfp.hMf = System.BitConverter.ToInt16(bytes, 8);

            if (this.mfp.mm > 98)
            {
                this.rcWinMf = new byte[14];
                Array.Copy(bytes, 10, this.rcWinMf, 0, 14);

                //dimensions
                this.dxaGoal = System.BitConverter.ToInt16(bytes, 24);
                this.dyaGoal = System.BitConverter.ToInt16(bytes, 26);
                this.mx = System.BitConverter.ToUInt16(bytes, 28);
                this.my = System.BitConverter.ToUInt16(bytes, 30);

                //cropping
                this.dxaCropLeft = System.BitConverter.ToInt16(bytes, 32);
                this.dyaCropTop = System.BitConverter.ToInt16(bytes, 34);
                this.dxaCropRight = System.BitConverter.ToInt16(bytes, 36);
                this.dyaCropBottom = System.BitConverter.ToInt16(bytes, 38);

                //border
                byte[] brc = new byte[4];
                Array.Copy(bytes, 46, brc, 0, 4);
                this.brcTop = new BorderCode(brc);
                Array.Copy(bytes, 50, brc, 0, 4);
                this.brcLeft = new BorderCode(brc);
                Array.Copy(bytes, 54, brc, 0, 4);
                this.brcBottom = new BorderCode(brc);
                Array.Copy(bytes, 58, brc, 0, 4);
                this.brcRight = new BorderCode(brc);

                this.dxaOrigin = System.BitConverter.ToInt16(bytes, 62);
                this.dyaOrigin = System.BitConverter.ToInt16(bytes, 64);
                this.cProps = System.BitConverter.ToInt16(bytes, 66);

                //read the picture
                Int32 picStart = getPictureStart(bytes);
                if (picStart != -1)
                {
                    this.Picture = new byte[bytes.Length - picStart];
                    Array.Copy(bytes, picStart, this.Picture, 0, this.Picture.Length);
                }
            }
        }

        /// <summary>
        /// Searches for the beginning of the real picture file.
        /// </summary>
        /// <param name="pic">the bytes of the PIC structure</param>
        /// <returns>the position where the picture starts</returns>
        private int getPictureStart(byte[] pic)
        {
            int ret = -1;

            for (int i = 66; i < pic.Length; i++)
            {
                if (pic[i] == 0x89 && pic.Length > i + 3)
                {
                    //possible PNG start
                    if (pic[i + 1] == 0x50 && pic[i + 2] == 0x4E && pic[i + 3] == 0x47)
                    {
                        ret = i;
                        this.Type = PictureType.png;
                        break;
                    }
                }
                else if (pic[i] == 0xFF && pic.Length > i + 2)
                {
                    //possible PNG start
                    if (pic[i + 1] == 0xD8 && pic[i + 2] == 0xFF)
                    {
                        ret = i;
                        this.Type = PictureType.jpg;
                        break;
                    }
                }
            }

            return ret;
        }

        /// <summary>
        /// Returns the fcPic into the "data" stream, where the PIC begins.
        /// Returns -1 if the CHPX has no fcPic.
        /// </summary>
        /// <param name="chpx">The CHPX</param>
        /// <returns></returns>
        private Int32 getFcPic(CharacterPropertyExceptions chpx)
        {
            Int32 ret = -1;
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                switch (sprm.OpCode)
	            {
                    case 0x6A03:
                        ret = System.BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case 0x6a12:
                        ret = System.BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case 0x0806:
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
