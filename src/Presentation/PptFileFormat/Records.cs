using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.Reflection;
using System.Collections;

namespace PptFileFormat.Records
{
    public class Record
    {
        public const uint HEADER_SIZE_IN_BYTES = (16 + 16 + 32) / 8;

        public uint TotalSize
        {
            get { return HeaderSize + BodySize; }
        }

        public uint HeaderSize = HEADER_SIZE_IN_BYTES;
        public uint BodySize;

        public uint Type;
        public uint Version;
        public uint Instance;

        public Record(VirtualStream source, uint bodySize, uint type, uint version, uint instance)
        {
            this.BodySize = bodySize;
            this.Type = type;
            this.Version = version;
            this.Instance = instance;
        }

        public virtual string ToString(uint depth)
        {
            return String.Format("{0}{2}:\n{1}type = {3}, version = {4}, instance = {5}, bodySize = {6}",
                IndentationForDepth(depth), IndentationForDepth(depth + 1),
                this.GetType(), this.Type, this.Version, this.Instance, this.BodySize
            );
        }

        public override string ToString()
        {
            return this.ToString(0);
        }

        #region Static attributes and methods

        internal static string IndentationForDepth(uint depth)
        {
            StringBuilder result = new StringBuilder();

            for (uint i = 0; i < depth; i++)
                result.Append("  ");

            return result.ToString();
        }

        public static Dictionary<UInt16, Type> TypeToRecordClassMapping = GetTypeToRecordClassMapping();

        private static Dictionary<UInt16, Type> GetTypeToRecordClassMapping()
        {
            Dictionary<UInt16, Type> result = new Dictionary<UInt16, Type>();

            result.Add(1000, typeof(PptDocumentRecord));
            result.Add(1001, typeof(DocumentAtom));
            result.Add(1035, typeof(PPDrawingGroup));

            return result;
        }

        public static Record readRecord(VirtualStream source)
        {
            UInt16 verAndInstance = source.ReadUInt16();
            uint version = verAndInstance & 0x000FU;         // first 4 bit of field verAndInstance
            uint instance = (verAndInstance & 0xFFF0U) >> 4; // last 12 bit of field verAndInstance

            UInt16 type = source.ReadUInt16();
            UInt32 size = source.ReadUInt32();

            bool isContainer = (version == 0xF);

            Record result;
            Type cls;

            if (TypeToRecordClassMapping.TryGetValue(type, out cls))
            {
                ConstructorInfo constructor = cls.GetConstructor(new Type[] {
                    typeof(VirtualStream), typeof(uint), typeof(uint), typeof(uint), typeof(uint) 
                });

                if (constructor == null)
                {
                    throw new Exception(String.Format(
                        "Internal error: Could not find a matching constructor for class {0}",
                        cls));
                }

                try
                {
                    result = (Record)constructor.Invoke(new object[] {
                        source, size, type, version, instance
                    });
                }
                catch (TargetInvocationException e)
                {
                    System.Console.WriteLine(e.InnerException);
                    throw e.InnerException;
                }
            }
            else
            {
                result = new UnknownRecord(source, size, type, version, instance);
            }

            return result;
        }

        #endregion
    }

    public class UnknownRecord : Record
    {
        public UnknownRecord(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance)
        {
            source.Skip(this.BodySize);
        }
    }

    /// <summary>
    /// Regular containers are containers with Record children.
    /// 
    /// (There also is containers that only have a zipped XML payload.
    /// </summary>
    public class RegularContainer : Record
    {
        public List<Record> Children = new List<Record>();

        public RegularContainer(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance)
        {
            uint readSize = 0;
            while (readSize < this.BodySize)
            {
                Record child = Record.readRecord(source);

                this.Children.Add(child);
                readSize += child.TotalSize;
            }
        }

        override public string ToString(uint depth)
        {
            StringBuilder result = new StringBuilder(base.ToString(depth));
            result.Append("\n");

            depth++;
            result.Append(IndentationForDepth(depth));

            if (this.Children.Count > 0)
                result.Append("Children:");

            foreach (Record record in this.Children)
            {
                result.Append("\n");
                result.Append(record.ToString(depth + 1));
            }

            return result.ToString();
        }
    }

    public class PptDocumentRecord : RegularContainer
    {
        public PptDocumentRecord(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class DocumentAtom : Record
    {
        public GPointAtom SlideSize;
        public GPointAtom NotesSize;
        public GRatioAtom ServerZoom;

        public UInt32 NotesMasterPersist;
        public UInt32 HandoutMasterPersist;
        public UInt16 FirstSlideNum;
        public Int16 SlideSizeType;

        public bool SaveWithFonts;
        public bool OmitTitlePlace;
        public bool RightToLeft;
        public bool ShowComments;

        public DocumentAtom(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance)
        {
            this.SlideSize = new GPointAtom(source);
            this.NotesSize = new GPointAtom(source);
            this.ServerZoom = new GRatioAtom(source);

            this.NotesMasterPersist = source.ReadUInt32();
            this.HandoutMasterPersist = source.ReadUInt32();
            this.FirstSlideNum = source.ReadUInt16();
            this.SlideSizeType = source.ReadInt16();

            this.SaveWithFonts = source.ReadByte() != 0;
            this.OmitTitlePlace = source.ReadByte() != 0;
            this.RightToLeft = source.ReadByte() != 0;
            this.ShowComments = source.ReadByte() != 0;
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}slideSize = {2}, notesSize = {3}, serverZoom = {4}\n{1}" +
                "notesMasterPersist = {5}, handoutMasterPersist = {6}, firstSlideNum = {7}, slideSizeType = {8}\n{1}" +
                "saveWithFonts = {9}, omitTitlePlace = {10}, rightToLeft = {11}, showComments = {12}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.SlideSize, this.NotesSize, this.ServerZoom,

                this.NotesMasterPersist, this.HandoutMasterPersist, this.FirstSlideNum, this.SlideSizeType,

                this.SaveWithFonts, this.OmitTitlePlace, this.RightToLeft, this.ShowComments);
        }
    }

    public class GPointAtom
    {
        public Int32 x;
        public Int32 y;

        public GPointAtom(VirtualStream source)
        {
            this.x = source.ReadInt32();
            this.y = source.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("PointAtom({0}, {1})", this.x, this.y);
        }
    }

    public class GRatioAtom
    {
        public Int32 numer;
        public Int32 denom;

        public GRatioAtom(VirtualStream source)
        {
            this.numer = source.ReadInt32();
            this.denom = source.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("RatioAtom({0}, {1})", this.numer, this.denom);
        }
    }

    public class PPDrawingGroup : RegularContainer
    {
        public PPDrawingGroup(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }
}
