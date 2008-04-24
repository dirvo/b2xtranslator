using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(TypeCode = 4001)]
    public class StyleTextPropAtom : Record
    {
        [FlagsAttribute]
        public enum ParagraphMask : uint
        {
            None = 0,
            HasCustomBullet = 1 << 0,
            HasCustomBulletTypeface = 1 << 1,
            HasCustomBulletColor = 1 << 2,
            HasCustomBulletSize = 1 << 3,

            BulletFlagsFieldPresent = HasCustomBullet | HasCustomBulletTypeface |
                                          HasCustomBulletColor | HasCustomBulletSize,

            BulletTypefacePresent = 1 << 4,
            BulletSizePresent = 1 << 5,
            BulletColorPresent = 1 << 6,
            BulletCharPresent = 1 << 7,
            LeftMarginPresent = 1 << 8,

            // Bit 9 is unused

            IndentPresent = 1 << 10,
            AlignmentPresent = 1 << 11,
            LineSpacingPresent = 1 << 12,
            SpaceBeforePresent = 1 << 13,
            SpaceAfterPresent = 1 << 14,
            DefaultTabSizePresent = 1 << 15,
            BaseLinePresent = 1 << 16,

            HasCustomCharWrap = 1 << 17,
            HasCustomWordWrap = 1 << 18,
            HasCustomOverflow = 1 << 19,

            LineBreakFlagsFieldPresent = HasCustomCharWrap | HasCustomWordWrap | HasCustomOverflow,

            TabStopsPresent = 1 << 20,
            TextDirectionPresent = 1 << 21
        }

        public class ParagraphRun
        {
            public UInt32 Length;
            public UInt16 IndentLevel;
            public ParagraphMask Mask;

            #region Presence flag getters
            public bool BulletFlagsFieldPresent
            {
                get { return (this.Mask & ParagraphMask.BulletFlagsFieldPresent) != 0; }
            }

            public bool BulletCharPresent
            {
                get { return (this.Mask & ParagraphMask.BulletCharPresent) != 0; }
            }

            public bool BulletTypefacePresent
            {
                get { return (this.Mask & ParagraphMask.BulletTypefacePresent) != 0; }
            }

            public bool BulletSizePresent
            {
                get { return (this.Mask & ParagraphMask.BulletSizePresent) != 0; }
            }

            public bool BulletColorPresent
            {
                get { return (this.Mask & ParagraphMask.BulletColorPresent) != 0; }
            }

            public bool AlignmentPresent
            {
                get { return (this.Mask & ParagraphMask.AlignmentPresent) != 0; }
            }

            public bool LineSpacingPresent
            {
                get { return (this.Mask & ParagraphMask.LineSpacingPresent) != 0; }
            }

            public bool SpaceBeforePresent
            {
                get { return (this.Mask & ParagraphMask.SpaceBeforePresent) != 0; }
            }

            public bool SpaceAfterPresent
            {
                get { return (this.Mask & ParagraphMask.SpaceAfterPresent) != 0; }
            }

            public bool LeftMarginPresent
            {
                get { return (this.Mask & ParagraphMask.LeftMarginPresent) != 0; }
            }

            public bool IndentPresent
            {
                get { return (this.Mask & ParagraphMask.IndentPresent) != 0; }
            }

            public bool DefaultTabSizePresent
            {
                get { return (this.Mask & ParagraphMask.DefaultTabSizePresent) != 0; }
            }

            public bool TabStopsPresent
            {
                get { return (this.Mask & ParagraphMask.TabStopsPresent) != 0; }
            }

            public bool BaseLinePresent
            {
                get { return (this.Mask & ParagraphMask.BaseLinePresent) != 0; }
            }

            public bool LineBreakFlagsFieldPresent
            {
                get { return (this.Mask & ParagraphMask.LineBreakFlagsFieldPresent) != 0; }
            }

            public bool TextDirectionPresent
            {
                get { return (this.Mask & ParagraphMask.TextDirectionPresent) != 0; }
            }
            #endregion

            public UInt16? BulletFlags;
            public char? BulletChar;
            public UInt16? BulletTypefaceIdx;
            public Int16? BulletSize;
            public GrColorAtom BulletColor;
            public Int16? Alignment;
            public Int16? LineSpacing;
            public Int16? SpaceBefore;
            public Int16? SpaceAfter;
            public Int16? LeftMargin;
            public Int16? Indent;
            public Int16? DefaultTabSize;
            public UInt16? BaseLine;
            public UInt16? LineBreakFlags;
            public UInt16? TextDirection;

            public ParagraphRun(BinaryReader reader)
            {
                this.Length = reader.ReadUInt32();
                this.IndentLevel = reader.ReadUInt16();
                this.Mask = (ParagraphMask)reader.ReadUInt32();

                // Note: These appear in Mask as well -- there they are true
                // when the flag differs from the Master style.
                // The actual value for the differing flags is stored here.
                // (TODO: This is still a guess. Verify.)
                if (this.BulletFlagsFieldPresent)
                    this.BulletFlags = reader.ReadUInt16();

                if (this.BulletCharPresent)
                    this.BulletChar = (char)reader.ReadUInt16();

                if (this.BulletTypefacePresent)
                    this.BulletTypefaceIdx = reader.ReadUInt16();

                if (this.BulletSizePresent)
                    this.BulletSize = reader.ReadInt16();

                if (this.BulletColorPresent)
                    this.BulletColor = new GrColorAtom(reader);

                if (this.AlignmentPresent)
                    this.Alignment = reader.ReadInt16();

                if (this.LineSpacingPresent)
                    this.LineSpacing = reader.ReadInt16();

                if (this.SpaceBeforePresent)
                    this.SpaceBefore = reader.ReadInt16();

                if (this.SpaceAfterPresent)
                    this.SpaceAfter = reader.ReadInt16();

                if (this.LeftMarginPresent)
                    this.LeftMargin = reader.ReadInt16();

                if (this.IndentPresent)
                    this.Indent = reader.ReadInt16();

                if (this.DefaultTabSizePresent)
                    this.DefaultTabSize = reader.ReadInt16();

                if (this.TabStopsPresent)
                {
                    UInt16 tabStopsCount = reader.ReadUInt16();
                    if (tabStopsCount != 0)
                        throw new NotImplementedException("Tab stop reading not yet implemented"); // TODO
                }

                if (this.BaseLinePresent)
                    this.BaseLine = reader.ReadUInt16();

                if (this.LineBreakFlagsFieldPresent)
                    this.LineBreakFlags = reader.ReadUInt16();

                if (this.TextDirectionPresent)
                    this.TextDirection = reader.ReadUInt16();
            }

            public string ToString(uint depth)
            {
                StringBuilder result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.Append(base.ToString());

                depth++;
                string indent = IndentationForDepth(depth);

                result.AppendFormat("\n{0}Length = {1}", indent, this.Length);
                result.AppendFormat("\n{0}IndentLevel = {1}", indent, this.IndentLevel);
                result.AppendFormat("\n{0}Mask = {1}", indent, this.Mask);

                if (this.BulletFlagsFieldPresent)
                    result.AppendFormat("\n{0}BulletFlags = {1}", indent, this.BulletFlags);

                if (this.BulletCharPresent)
                    result.AppendFormat("\n{0}BulletChar = {1}", indent, this.BulletChar);

                if (this.BulletTypefacePresent)
                    result.AppendFormat("\n{0}BulletTypefaceIdx = {1}", indent, this.BulletTypefaceIdx);

                if (this.BulletSizePresent)
                    result.AppendFormat("\n{0}BulletSize = {1}", indent, this.BulletSize);

                if (this.BulletColorPresent)
                    result.AppendFormat("\n{0}BulletColor = {1}", indent, this.BulletColor);

                if (this.AlignmentPresent)
                    result.AppendFormat("\n{0}Alignment = {1}", indent, this.Alignment);

                if (this.LineSpacingPresent)
                    result.AppendFormat("\n{0}LineSpacing = {1}", indent, this.LineSpacing);

                if (this.SpaceBeforePresent)
                    result.AppendFormat("\n{0}SpaceBefore = {1}", indent, this.SpaceBefore);

                if (this.SpaceAfterPresent)
                    result.AppendFormat("\n{0}SpaceAfter = {1}", indent, this.SpaceAfter);

                if (this.LeftMarginPresent)
                    result.AppendFormat("\n{0}LeftMargin = {1}", indent, this.LeftMargin);

                if (this.IndentPresent)
                    result.AppendFormat("\n{0}Indent = {1}", indent, this.Indent);

                if (this.DefaultTabSizePresent)
                    result.AppendFormat("\n{0}DefaultTabSize = {1}", indent, this.DefaultTabSize);

                if (this.BaseLinePresent)
                    result.AppendFormat("\n{0}BaseLine = {1}", indent, this.BaseLine);

                if (this.LineBreakFlagsFieldPresent)
                    result.AppendFormat("\n{0}LineBreakFlags = {1}", indent, this.LineBreakFlags);

                if (this.TextDirectionPresent)
                    result.AppendFormat("\n{0}TextDirection = {1}", indent, this.TextDirection);

                return result.ToString();
            }

            public override string ToString()
            {
                return this.ToString(0);
            }
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
            SymbolTypefacePresent = 1 << 23
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
            HasAsianSmartQuotes = 1 << 5,

            // Bit 6 is unused

            HasHorizonNumRendering = 1 << 7,

            // Bit 8 is unused

            IsEmbossed = 1 << 9,

            ExtensionNibble = 0xF << 10 // Bit 10 - 13

            // Bit 14 - 15 are unused
        }

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
                this.Length = reader.ReadUInt32();
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

                if (this.PositionPresent)
                    this.Position = reader.ReadUInt16();

                if (this.ColorPresent)
                    this.Color = new GrColorAtom(reader);
            }

            public string ToString(uint depth)
            {
                StringBuilder result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.Append(base.ToString());

                depth++;
                string indent = IndentationForDepth(depth);

                result.AppendFormat("\n{0}Length = {1}", indent, this.Length);
                result.AppendFormat("\n{0}Mask = {1}", indent, this.Mask);

                if (this.StyleFlagsFieldPresent)
                    result.AppendFormat("\n{0}Style = {1}", indent, this.Style);

                if (this.TypefacePresent)
                    result.AppendFormat("\n{0}TypefaceIdx = {1}", indent, this.TypefaceIdx);

                if (this.FEOldTypefacePresent)
                    result.AppendFormat("\n{0}FEOldTypefaceIdx = {1}", indent, this.FEOldTypefaceIdx);

                if (this.ANSITypefacePresent)
                    result.AppendFormat("\n{0}ANSITypefaceIdx = {1}", indent, this.ANSITypefaceIdx);

                if (this.SymbolTypefacePresent)
                    result.AppendFormat("\n{0}SymbolTypefaceIdx = {1}", indent, this.SymbolTypefaceIdx);

                if (this.SizePresent)
                    result.AppendFormat("\n{0}Size = {1}", indent, this.Size);

                if (this.PositionPresent)
                    result.AppendFormat("\n{0}Position = {1}", indent, this.Position);

                if (this.ColorPresent)
                    result.AppendFormat("\n{0}Color = {1}", indent, this.Color);

                return result.ToString();
            }

            public override string ToString()
            {
                return this.ToString(0);
            }
        }

        public StyleTextPropAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
        }

        override public void AfterParentSet()
        {
            // Anmerkung: In OOXML kann ein Character-Properties-Element sich
            // nicht über mehrere Paragraphen hinweg erstrecken.
            // TODO: War das im Binärformat der Fall?

            // TODO: FindParentByType? FindChildByType? FindSiblingByType?

            ClientTextbox textbox = this.ParentRecord as ClientTextbox;
            if (textbox == null)
                throw new Exception("Record of type StyleTextPropAtom doesn't have parent of type ClientTextbox");

            TextAtom textAtom = textbox.Children.Find(
                delegate(Record sibling) { return sibling is TextAtom; }
            ) as TextAtom;

            /* This can legitimately happen... */
            if (textAtom == null)
            {
                this.Reader.Read(new char[this.BodySize], 0, (int)this.BodySize);
                return;
            }

            // TODO: Length in bytes? UTF-16 characters? Full width unicode characters?

            uint seenLength = 0;
            while (seenLength < textAtom.Text.Length)
            {
                ParagraphRun run = new ParagraphRun(this.Reader);
                Console.WriteLine(run.ToString());
                Console.WriteLine("  Text = {0}", Utils.StringInspect(
                    textAtom.Text.Substring((int)seenLength, (int)run.Length)));
                Console.WriteLine();

                seenLength += run.Length;
            }

            seenLength = 0;
            while (seenLength < textAtom.Text.Length)
            {
                CharacterRun run = new CharacterRun(this.Reader);
                Console.WriteLine(run.ToString());
                Console.WriteLine("  Text = {0}", Utils.StringInspect(
                    textAtom.Text.Substring((int)seenLength, (int)run.Length)));
                Console.WriteLine();

                seenLength += run.Length;
            }
        }

        /*        public override string ToString(uint depth)
                {
                    return String.Format("{0}\n{1}RunLength = {2}",
                        base.ToString(depth), IndentationForDepth(depth + 1), this.RunLength);
                }*/
    }

}
