using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4116)]
    public class AnimationInfoContainer : RegularContainer
    {
        public AnimationInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(4081)]
    public class AnimationInfoAtom : Record
    {
        public byte[] dimColor;
        public Int16 flags;
        public byte[] soundIdRef;
        public Int32 delayTime;
        public Int16 orderID;
        public UInt16 slideCount;
        public AnimBuildTypeEnum animBuildType;
        public byte animEffect;
        public byte animEffectDirection;
        public AnimAfterEffectEnum animAfterEffect;
        public TextBuildSubEffectEnum textBuildSubEffect;
        public byte oleVerb;

        public bool fReverse;
        public bool fAutomatic;
        public bool fSound;
        public bool fStopSound;
        public bool fPlay;
        public bool fSynchronous;
        public bool fHide;
        public bool fAnimateBg;

        public AnimationInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

            dimColor = this.Reader.ReadBytes(4);
            flags = this.Reader.ReadInt16();

            fReverse = Tools.Utils.BitmaskToBool(flags, 0x1 << 0);
            fAutomatic = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fSound = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);
            fStopSound = Tools.Utils.BitmaskToBool(flags, 0x1 << 6);
            fPlay = Tools.Utils.BitmaskToBool(flags, 0x1 << 8);
            fSynchronous = Tools.Utils.BitmaskToBool(flags, 0x1 << 10);
            fHide = Tools.Utils.BitmaskToBool(flags, 0x1 << 12);
            fAnimateBg = Tools.Utils.BitmaskToBool(flags, 0x1 << 14);

            Int16 reserved = this.Reader.ReadInt16();
            soundIdRef = this.Reader.ReadBytes(4);
            delayTime = this.Reader.ReadInt32();
            orderID = this.Reader.ReadInt16();
            slideCount = this.Reader.ReadUInt16();
            animBuildType = (AnimBuildTypeEnum)this.Reader.ReadByte();
            animEffect = this.Reader.ReadByte();
            animEffectDirection = this.Reader.ReadByte();
            animAfterEffect = (AnimAfterEffectEnum)this.Reader.ReadByte();
            textBuildSubEffect = (TextBuildSubEffectEnum)this.Reader.ReadByte();
            oleVerb = this.Reader.ReadByte();

            if (this.Reader.BaseStream.Position != this.Reader.BaseStream.Length)
            {
                this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
            }
        }

        [FlagsAttribute]
        public enum AnimationFlagsMask : uint
        {
            None = 0,

            fReverse = 3 << 14,
            fAutomatic = 3 << 12,
            fSound = 3 << 10,
            fStopSound = 3 << 8,
            fPlay = 3 << 6,
            fSynchronous = 3 << 4,
            fHide = 3 << 2,
            fAnimateBg = 3
        }

        public enum AnimBuildTypeEnum: byte
        {
            FollowMaster = 0xFE,

            NoBuild = 0x00,
            OneBuild = 0x01,
            Level1Build = 0x02,
            Level2Build = 0x03,
            Level3Build = 0x04,
            Level4Build = 0x05,
            Level5Build = 0x06,
            GraphBySeries = 0x07,
            GraphByCategory = 0x08,
            GraphByElementInSeries = 0x09,
            GraphByElementInCategory = 0x0A
        }

        public enum AnimAfterEffectEnum : byte
        {
            NoAfterEffect = 0x00,
            Dim = 0x01,
            Hide = 0x02,
            HideImmediately = 0x03
        }

        public enum TextBuildSubEffectEnum : byte
        {
            BuildByNone = 0x00,
            BuildByWord = 0x01,
            BuildByCharacter = 0x02
        }
    }

}
