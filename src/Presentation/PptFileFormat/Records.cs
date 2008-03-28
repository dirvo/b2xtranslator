using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.Reflection;
using System.Collections;
using System.IO;

namespace PptFileFormat.Records
{
    public class Record : IEnumerable<Record>
    {
        public const uint HEADER_SIZE_IN_BYTES = (16 + 16 + 32) / 8;

        public uint TotalSize
        {
            get { return HeaderSize + BodySize; }
        }

        public Record ParentRecord = null;

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

        public string GetIdentifier()
        {
            StringBuilder result = new StringBuilder();

            Record r = this;
            bool isFirst = true;

            while (r != null)
            {
                if (!isFirst)
                    result.Insert(0, "-");

                result.Insert(0, String.Format("{0}i{1}", r.FormatType(), r.Instance));

                r = r.ParentRecord;
                isFirst = false;
            }

            return result.ToString();
        }

        public string FormatType()
        {
            bool isEscherRecord = (this.Type >= 0xF000 && this.Type <= 0xFFFF);
            return String.Format(isEscherRecord ? "0x{0:X}" : "{0}", this.Type);
        }

        public virtual string ToString(uint depth)
        {
            return String.Format("{0}{2}:\n{1}Type = {3}, Version = {4}, Instance = {5}, BodySize = {6}",
                IndentationForDepth(depth), IndentationForDepth(depth + 1),
                this.GetType(), this.FormatType(), this.Version, this.Instance, this.BodySize
            );
        }

        public override string ToString()
        {
            return this.ToString(0);
        }

        #region IEnumerable<Record> Members

        public virtual IEnumerator<Record> GetEnumerator()
        {
            yield return this;
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (Record record in this)
                yield return record;
        }

        #endregion

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

            // PowerPoint records
            result.Add(1000, typeof(PptDocumentRecord));
            result.Add(1001, typeof(DocumentAtom));
            result.Add(1006, typeof(Slide));
            result.Add(1016, typeof(List));
            result.Add(1035, typeof(PPDrawingGroup));
            result.Add(1036, typeof(PPDrawing));
            result.Add(4008, typeof(TextBytesAtom));
            result.Add(4080, typeof(SlideListWithText));

            // Drawing records
            result.Add(0xF000, typeof(DrawingGroup));
            result.Add(0xF002, typeof(DrawingContainer));
            result.Add(0xF003, typeof(GroupContainer));
            result.Add(0xF004, typeof(ShapeContainer));
            result.Add(0xF006, typeof(DrawingGroupRecord));
            result.Add(0xF00D, typeof(ClientTextbox));
            result.Add(0xF011, typeof(ClientData));

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
        public byte[] RawData;

        public UnknownRecord(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance)
        {
            this.RawData = new byte[this.BodySize];
            source.Read(this.RawData);
        }

        public void DumpToStream(Stream output)
        {
            using (BinaryWriter writer = new BinaryWriter(output))
            {
                writer.Write(this.RawData, 0, this.RawData.Length);
            }
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
                child.ParentRecord = this;

                readSize += child.TotalSize;
            }
        }

        override public string ToString(uint depth)
        {
            StringBuilder result = new StringBuilder(base.ToString(depth));

            depth++;

            if (this.Children.Count > 0)
            {
                result.AppendLine();
                result.Append(IndentationForDepth(depth));
                result.Append("Children:");
            }

            foreach (Record record in this.Children)
            {
                result.AppendLine();
                result.Append(record.ToString(depth + 1));
            }

            return result.ToString();
        }

        #region IEnumerable<Record> Members

        public override IEnumerator<Record> GetEnumerator()
        {
            yield return this;

            foreach (Record recordChild in this.Children)
                foreach (Record record in recordChild)
                    yield return record;
        }

        #endregion
    }

    #region PowerPoint records
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
            return String.Format("{0}\n{1}SlideSize = {2}, NotesSize = {3}, ServerZoom = {4}\n{1}" +
                "NotesMasterPersist = {5}, HandoutMasterPersist = {6}, FirstSlideNum = {7}, SlideSizeType = {8}\n{1}" +
                "SaveWithFonts = {9}, OmitTitlePlace = {10}, RightToLeft = {11}, ShowComments = {12}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.SlideSize, this.NotesSize, this.ServerZoom,

                this.NotesMasterPersist, this.HandoutMasterPersist, this.FirstSlideNum, this.SlideSizeType,

                this.SaveWithFonts, this.OmitTitlePlace, this.RightToLeft, this.ShowComments);
        }
    }

    public class GPointAtom
    {
        public Int32 X;
        public Int32 Y;

        public GPointAtom(VirtualStream source)
        {
            this.X = source.ReadInt32();
            this.Y = source.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("PointAtom({0}, {1})", this.X, this.Y);
        }
    }

    public class GRatioAtom
    {
        public Int32 Numer;
        public Int32 Denom;

        public GRatioAtom(VirtualStream source)
        {
            this.Numer = source.ReadInt32();
            this.Denom = source.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("RatioAtom({0}, {1})", this.Numer, this.Denom);
        }
    }

    public class PPDrawing : RegularContainer
    {
        public PPDrawing(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class PPDrawingGroup : RegularContainer
    {
        public PPDrawingGroup(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class SlideListWithText : RegularContainer
    {
        public SlideListWithText(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class List : RegularContainer
    {
        public List(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class Slide : RegularContainer
    {
        public Slide(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class TextBytesAtom : Record
    {
        public string Text;

        public TextBytesAtom(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) {
            byte[] bytes = new byte[size];
            source.Read(bytes, (int) size);

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < size; i++)
                sb.Append((char) bytes[i]);

            this.Text = sb.ToString();
        }

        public override string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Text = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1),  this.Text);
        }
    }

    #endregion

    #region Drawing records
    public class DrawingGroup : RegularContainer
    {
        public DrawingGroup(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class DrawingGroupRecord : Record
    {
        public class FileIdCluster
        {
            public UInt32 DrawingGroupId;
            public UInt32 CSpIdCur;

            public FileIdCluster(VirtualStream source)
            {
                this.DrawingGroupId = source.ReadUInt32();
                this.CSpIdCur = source.ReadUInt32();
            }

            public string ToString(uint depth)
            {
                StringBuilder result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.AppendFormat("FileIdCluster: DrawingGroupId = {0}, CSpIdCur = {1}",
                   this.DrawingGroupId, this.CSpIdCur);

                return result.ToString();
            }
        }

        public UInt32 MaxShapeId;           // Maximum shape ID
        public UInt32 IdClustersCount;      // Number of FileIdClusters
        public UInt32 ShapesSavedCount;     // Total number of shapes saved
        public UInt32 DrawingsSavedCount;   // Total number of drawings saved

        public List<FileIdCluster> Clusters = new List<FileIdCluster>();

        public DrawingGroupRecord(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance)
        {
            this.MaxShapeId = source.ReadUInt32();
            this.IdClustersCount = source.ReadUInt32() - 1; // Office saves the actual value + 1 -- flgr
            this.ShapesSavedCount = source.ReadUInt32();
            this.DrawingsSavedCount = source.ReadUInt32();

            for (int i = 0; i < this.IdClustersCount; i++)
            {
                Clusters.Add(new FileIdCluster(source));
            }
        }

        override public string ToString(uint depth)
        {
            StringBuilder result = new StringBuilder();

            result.AppendLine(base.ToString(depth));

            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("MaxShapeId = {0}, IdClustersCount = {1}",
                this.MaxShapeId, this.IdClustersCount);
                
            result.AppendLine();
            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("ShapesSavedCount = {0}, DrawingsSavedCount = {1}",
                this.ShapesSavedCount, this.DrawingsSavedCount);

            depth++;

            if (this.Clusters.Count > 0)
            {
                result.AppendLine();
                result.Append(IndentationForDepth(depth));
                result.Append("Clusters:");
            }

            foreach (FileIdCluster cluster in this.Clusters)
            {
                result.AppendLine();
                result.Append(cluster.ToString(depth + 1));
            }

            return result.ToString();
        }
    }

    public class DrawingContainer : RegularContainer
    {
        public DrawingContainer(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class GroupContainer : RegularContainer
    {
        public GroupContainer(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class ShapeContainer : RegularContainer
    {
        public ShapeContainer(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class ClientTextbox : RegularContainer
    {
        public ClientTextbox(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    public class ClientData : RegularContainer
    {
        public ClientData(VirtualStream source, uint size, uint type, uint version, uint instance)
            : base(source, size, type, version, instance) { }
    }

    #endregion
}
