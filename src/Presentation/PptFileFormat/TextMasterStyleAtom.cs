using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    //[OfficeRecordAttribute(4003)]
    class TextMasterStyleAtom : TextStyleAtom
    {
        public UInt16 IndentLevelCount;

        public TextMasterStyleAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.IndentLevelCount = this.Reader.ReadUInt16();

            for (int i = 0; i < this.IndentLevelCount; i++)
            {
                long pos = this.Reader.BaseStream.Position;

                this.PRuns[i] = new ParagraphRun(this.Reader, true);

                TraceLogger.DebugInternal("Read paragraph run. Before pos = {0}, after pos = {1} of {2}: {3}",
                    pos, this.Reader.BaseStream.Position, this.Reader.BaseStream.Length,
                    this.PRuns[i]);

                pos = this.Reader.BaseStream.Position;
                this.CRuns[i] = new CharacterRun(this.Reader);

                TraceLogger.DebugInternal("Read character run. Before pos = {0}, after pos = {1} of {2}: {3}",
                    pos, this.Reader.BaseStream.Position, this.Reader.BaseStream.Length,
                    this.CRuns[i]);
            }

            // XXX: I'm not sure why but in some cases there is trailing garbage -- flgr
            this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
        }

        public ParagraphRun ParagraphRunForIndentLevel(int level)
        {
            return this.PRuns[level];
        }

        public CharacterRun CharacterRunForIndentLevel(int level)
        {
            return this.CRuns[level];
        }
    }
}
