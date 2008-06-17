using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4003)]
    class TxMasterStyleAtom : Record
    {
        public UInt16 IndentLevelCount;

        private ParagraphRun[] pruns;
        private CharacterRun[] cruns;

        public TxMasterStyleAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.IndentLevelCount = this.Reader.ReadUInt16();

            this.pruns = new ParagraphRun[this.IndentLevelCount];
            this.cruns = new CharacterRun[this.IndentLevelCount];

            for (int i = 0; i < this.IndentLevelCount; i++)
            {
                this.pruns[i] = new ParagraphRun(this.Reader, true);
                this.cruns[i] = new CharacterRun(this.Reader);
            }

            // XXX: I'm not sure why but in some cases there is trailing garbage -- flgr
            this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
        }

        public ParagraphRun ParagraphRunForIndentLevel(int level)
        {
            return this.pruns[level];
        }

        public CharacterRun CharacterRunForIndentLevel(int level)
        {
            return this.cruns[level];
        }

        public override string ToString(uint depth)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(base.ToString(depth));

            depth++;
            string indent = IndentationForDepth(depth);

            sb.AppendFormat("\n{0}Paragraph Runs:", indent);
            foreach (ParagraphRun pr in this.pruns)
                sb.AppendFormat("\n{0}", pr.ToString(depth + 1));

            sb.AppendFormat("\n{0}Character Runs:", indent);
            foreach (CharacterRun cr in this.cruns)
                sb.AppendFormat("\n{0}", cr.ToString(depth + 1));

            return sb.ToString();
        }
    }
}
