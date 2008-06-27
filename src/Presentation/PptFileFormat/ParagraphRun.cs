using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class ParagraphRun
    {
        public class TabStop
        {
            public UInt16 Distance;
            public UInt16 Alignment;

            public TabStop(BinaryReader reader)
            {
                this.Distance = reader.ReadUInt16();
                this.Alignment = reader.ReadUInt16();
            }
        }

        public UInt32 Length;
        public UInt16 IndentLevel;
        public ParagraphMask Mask;

        public TabStop[] TabStops;

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

        public bool HasCustomBulletSize
        {
            get { return (this.Mask & ParagraphMask.HasCustomBulletSize) != 0; }
        }

        public bool BulletColorPresent
        {
            get { return (this.Mask & ParagraphMask.BulletColorPresent) != 0; }
        }

        public bool HasCustomBulletColor
        {
            get { return (this.Mask & ParagraphMask.HasCustomBulletColor) != 0; }
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

        public ParagraphRun(BinaryReader reader, bool noIndentField)
        {
            this.IndentLevel = noIndentField ? (ushort)0 : reader.ReadUInt16();
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

            if ((this.HasCustomBulletSize && (this.BulletFlags & (1 << 3)) != 0))
                this.BulletSize = reader.ReadInt16();

            if ((this.HasCustomBulletColor && (this.BulletFlags & (1 << 2)) != 0))
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

            if (this.BaseLinePresent)
                this.BaseLine = reader.ReadUInt16();

            if (this.LineBreakFlagsFieldPresent)
                this.LineBreakFlags = reader.ReadUInt16();

            if (this.TextDirectionPresent)
                this.TextDirection = reader.ReadUInt16();

            if (this.TabStopsPresent)
            {
                UInt16 tabStopsCount = reader.ReadUInt16();
                this.TabStops = new TabStop[tabStopsCount];

                for (int i = 0; i < tabStopsCount; i++)
                {
                    this.TabStops[i] = new TabStop(reader);
                }
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
            result.AppendFormat("\n{0}IndentLevel = {1}", indent, this.IndentLevel);
            result.AppendFormat("\n{0}Mask = {1}", indent, this.Mask);

            if (this.BulletFlags != null)
                result.AppendFormat("\n{0}BulletFlags = {1}", indent, this.BulletFlags);

            if (this.BulletChar != null)
                result.AppendFormat("\n{0}BulletChar = {1}", indent, this.BulletChar);

            if (this.BulletTypefaceIdx != null)
                result.AppendFormat("\n{0}BulletTypefaceIdx = {1}", indent, this.BulletTypefaceIdx);

            if (this.BulletSize != null)
                result.AppendFormat("\n{0}BulletSize = {1}", indent, this.BulletSize);

            if (this.BulletColor != null)
                result.AppendFormat("\n{0}BulletColor = {1}", indent, this.BulletColor);

            if (this.Alignment != null)
                result.AppendFormat("\n{0}Alignment = {1}", indent, this.Alignment);

            if (this.LineSpacing != null)
                result.AppendFormat("\n{0}LineSpacing = {1}", indent, this.LineSpacing);

            if (this.SpaceBefore != null)
                result.AppendFormat("\n{0}SpaceBefore = {1}", indent, this.SpaceBefore);

            if (this.SpaceAfter != null)
                result.AppendFormat("\n{0}SpaceAfter = {1}", indent, this.SpaceAfter);

            if (this.LeftMargin != null)
                result.AppendFormat("\n{0}LeftMargin = {1}", indent, this.LeftMargin);

            if (this.Indent != null)
                result.AppendFormat("\n{0}Indent = {1}", indent, this.Indent);

            if (this.DefaultTabSize != null)
                result.AppendFormat("\n{0}DefaultTabSize = {1}", indent, this.DefaultTabSize);

            if (this.BaseLine != null)
                result.AppendFormat("\n{0}BaseLine = {1}", indent, this.BaseLine);

            if (this.LineBreakFlags != null)
                result.AppendFormat("\n{0}LineBreakFlags = {1}", indent, this.LineBreakFlags);

            if (this.TextDirection != null)
                result.AppendFormat("\n{0}TextDirection = {1}", indent, this.TextDirection);

            return result.ToString();
        }

        public override string ToString()
        {
            return this.ToString(0);
        }
    }

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
}
