using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class CharacterRun
    {
        public UInt32 Length;
        public CharacterMask Mask;

        #region Presence flag getters
        public bool StyleFlagsFieldPresent
        {
            get { return (this.Mask & CharacterMask.StyleFlagsFieldPresent) != 0; }
        }

        public bool TypefacePresent
        {
            get { return (this.Mask & CharacterMask.TypefacePresent) != 0; }
        }

        public bool FEOldTypefacePresent
        {
            get { return (this.Mask & CharacterMask.FEOldTypefacePresent) != 0; }
        }

        public bool ANSITypefacePresent
        {
            get { return (this.Mask & CharacterMask.ANSITypefacePresent) != 0; }
        }

        public bool SymbolTypefacePresent
        {
            get { return (this.Mask & CharacterMask.SymbolTypefacePresent) != 0; }
        }

        public bool SizePresent
        {
            get { return (this.Mask & CharacterMask.SizePresent) != 0; }
        }

        public bool PositionPresent
        {
            get { return (this.Mask & CharacterMask.PositionPresent) != 0; }
        }

        public bool ColorPresent
        {
            get { return (this.Mask & CharacterMask.ColorPresent) != 0; }
        }
        #endregion

        public StyleMask? Style;
        public UInt16? TypefaceIdx;
        public UInt16? FEOldTypefaceIdx;
        public UInt16? ANSITypefaceIdx;
        public UInt16? SymbolTypefaceIdx;
        public UInt16? Size;
        public UInt16? Position;
        public GrColorAtom Color;

        public CharacterRun(BinaryReader reader)
        {
            try
            {
                this.Mask = (CharacterMask)reader.ReadUInt32();

                if (this.StyleFlagsFieldPresent)
                    this.Style = (StyleMask)reader.ReadUInt16();

                if (this.TypefacePresent)
                    this.TypefaceIdx = reader.ReadUInt16();

                if (this.FEOldTypefacePresent)
                    this.FEOldTypefaceIdx = reader.ReadUInt16();

                if (this.ANSITypefacePresent)
                    this.ANSITypefaceIdx = reader.ReadUInt16();

                if (this.SymbolTypefacePresent)
                    this.SymbolTypefaceIdx = reader.ReadUInt16();

                if (this.SizePresent)
                    this.Size = reader.ReadUInt16();

                if (this.ColorPresent)
                    this.Color = new GrColorAtom(reader);

                if (this.PositionPresent)
                    this.Position = reader.ReadUInt16();        
            }
            catch (EndOfStreamException e)
            {
                string s = e.ToString();
                //ignore
            }
               
        }

        public string ToString(uint depth)
        {
            StringBuilder result = new StringBuilder();

            string indent = Record.IndentationForDepth(depth);
            result.Append(indent);
            result.Append(base.ToString());

            depth++;
            indent = Record.IndentationForDepth(depth);

            result.AppendFormat("\n{0}Length = {1}", indent, this.Length);
            result.AppendFormat("\n{0}Mask = {1}", indent, this.Mask);

            if (this.Style != null)
                result.AppendFormat("\n{0}Style = {1}", indent, this.Style);

            if (this.TypefaceIdx != null)
                result.AppendFormat("\n{0}TypefaceIdx = {1}", indent, this.TypefaceIdx);

            if (this.FEOldTypefaceIdx != null)
                result.AppendFormat("\n{0}FEOldTypefaceIdx = {1}", indent, this.FEOldTypefaceIdx);

            if (this.ANSITypefaceIdx != null)
                result.AppendFormat("\n{0}ANSITypefaceIdx = {1}", indent, this.ANSITypefaceIdx);

            if (this.SymbolTypefaceIdx != null)
                result.AppendFormat("\n{0}SymbolTypefaceIdx = {1}", indent, this.SymbolTypefaceIdx);

            if (this.Size != null)
                result.AppendFormat("\n{0}Size = {1}", indent, this.Size);

            if (this.Position != null)
                result.AppendFormat("\n{0}Position = {1}", indent, this.Position);

            if (this.Color != null)
                result.AppendFormat("\n{0}Color = {1}", indent, this.Color);

            return result.ToString();
        }

        public override string ToString()
        {
            return this.ToString(0);
        }
    }

    [FlagsAttribute]
    public enum StyleMask : uint
    {
        None = 0,

        IsBold = 1 << 0,
        IsItalic = 1 << 1,
        IsUnderlined = 1 << 2,

        // Bit 3 is unused

        HasShadow = 1 << 4,
        HasAsianSmartQuotes = 1 << 5, //this should be fehint?

        // Bit 6 is unused

        HasHorizonNumRendering = 1 << 7, //this should be kumi?

        // Bit 8 is unused

        IsEmbossed = 1 << 9,

        ExtensionNibble = 0xF << 10 // Bit 10 - 13

        // Bit 14 - 15 are unused
    }

    [FlagsAttribute]
    public enum CharacterMask : uint
    {
        None = 0,

        // Bit 0 - 15 are used for marking style flag presence
        StyleFlagsFieldPresent = 0xFFFF,

        TypefacePresent = 1 << 16,
        SizePresent = 1 << 17,
        ColorPresent = 1 << 18,
        PositionPresent = 1 << 19,

        // Bit 20 is unused

        FEOldTypefacePresent = 1 << 21,
        ANSITypefacePresent = 1 << 22,
        SymbolTypefacePresent = 1 << 23,
        pp11ext = 1 << 26
    }
}
