using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.Reflection;
using System.Collections;
using System.IO;

namespace PptFileFormat.Records
{
    /// <summary>
    /// Used for mapping Office record TypeCodes to the classes implementing them.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class OfficeRecord : Attribute
    {
        public UInt16 TypeCode;
    }

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

        public byte[] RawData;

        protected BinaryReader Reader;

        public uint TypeCode;
        public uint Version;
        public uint Instance;

        public Record(BinaryReader _reader, uint bodySize, uint typeCode, uint version, uint instance)
        {
            this.BodySize = bodySize;
            this.TypeCode = typeCode;
            this.Version = version;
            this.Instance = instance;

            this.RawData = _reader.ReadBytes((int)this.BodySize);

            this.Reader = new BinaryReader(new MemoryStream(this.RawData));
        }

        public void DumpToStream(Stream output)
        {
            using (BinaryWriter writer = new BinaryWriter(output))
            {
                writer.Write(this.RawData, 0, this.RawData.Length);
            }
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
            bool isEscherRecord = (this.TypeCode >= 0xF000 && this.TypeCode <= 0xFFFF);
            return String.Format(isEscherRecord ? "0x{0:X}" : "{0}", this.TypeCode);
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

            // Note: We return a Dictionary that maps Office record TypeCodes to Office record classes.
            // We do this by querying all classes in the current assembly, filtering by namespace
            // PptFileFormat.Records and looking for attributes of type OfficeRecord.
            //
            // If in doubt see usage below.
            foreach (Type t in Assembly.GetExecutingAssembly().GetTypes())
            {
                if (t.Namespace == "PptFileFormat.Records")
                {
                    object[] attrs = t.GetCustomAttributes(typeof(OfficeRecord), false);

                    OfficeRecord attr = null;
                    
                    if (attrs.Length > 0)
                        attr = attrs[0] as OfficeRecord;

                    if (attr != null)
                    {
                        UInt16 typeCode = attr.TypeCode;

                        if (result.ContainsKey(typeCode))
                        {
                            throw new Exception(String.Format(
                                "Tried to register TypeCode {0} to {1}, but it is already registered to {2}",
                                typeCode, t, result[typeCode]));
                        }

                        result.Add(attr.TypeCode, t);
                    }
                }
            }

            return result;
        }

        public static Record readRecord(Stream stream)
        {
            return readRecord(new BinaryReader(stream));
        }

        public static Record readRecord(BinaryReader reader)
        {
            UInt16 verAndInstance = reader.ReadUInt16();
            uint version = verAndInstance & 0x000FU;         // first 4 bit of field verAndInstance
            uint instance = (verAndInstance & 0xFFF0U) >> 4; // last 12 bit of field verAndInstance

            UInt16 typeCode = reader.ReadUInt16();
            UInt32 size = reader.ReadUInt32();

            bool isContainer = (version == 0xF);

            Record result;
            Type cls;

            if (TypeToRecordClassMapping.TryGetValue(typeCode, out cls))
            {
                ConstructorInfo constructor = cls.GetConstructor(new Type[] {
                    typeof(BinaryReader), typeof(uint), typeof(uint), typeof(uint), typeof(uint) 
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
                        reader, size, typeCode, version, instance
                    });
                }
                catch (TargetInvocationException e)
                {
                    Console.WriteLine(e.InnerException);
                    throw e.InnerException;
                }
            }
            else
            {
                result = new UnknownRecord(reader, size, typeCode, version, instance);
            }

            return result;
        }

        #endregion
    }

    public class UnknownRecord : Record
    {
        public UnknownRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    /// <summary>
    /// Regular containers are containers with Record children.
    /// 
    /// (There also is containers that only have a zipped XML payload.
    /// </summary>
    public class RegularContainer : Record
    {
        public List<Record> Children = new List<Record>();

        public RegularContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            uint readSize = 0;

            while (readSize < this.BodySize)
            {
                Record child = Record.readRecord(this.Reader);

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
    [OfficeRecord(TypeCode = 1000)]
    public class PptDocumentRecord : RegularContainer
    {
        public PptDocumentRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 1001)]
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

        public DocumentAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.SlideSize = new GPointAtom(this.Reader);
            this.NotesSize = new GPointAtom(this.Reader);
            this.ServerZoom = new GRatioAtom(this.Reader);

            this.NotesMasterPersist = this.Reader.ReadUInt32();
            this.HandoutMasterPersist = this.Reader.ReadUInt32();
            this.FirstSlideNum = this.Reader.ReadUInt16();
            this.SlideSizeType = this.Reader.ReadInt16();

            this.SaveWithFonts = this.Reader.ReadByte() != 0;
            this.OmitTitlePlace = this.Reader.ReadByte() != 0;
            this.RightToLeft = this.Reader.ReadByte() != 0;
            this.ShowComments = this.Reader.ReadByte() != 0;
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

        public GPointAtom(BinaryReader reader)
        {
            this.X = reader.ReadInt32();
            this.Y = reader.ReadInt32();
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

        public GRatioAtom(BinaryReader reader)
        {
            this.Numer = reader.ReadInt32();
            this.Denom = reader.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("RatioAtom({0}, {1})", this.Numer, this.Denom);
        }
    }

    [OfficeRecord(TypeCode = 1006)]
    public class Slide : RegularContainer
    {
        public Slide(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 1016)]
    public class List : RegularContainer
    {
        public List(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }


    [OfficeRecord(TypeCode = 1035)]
    public class PPDrawingGroup : RegularContainer
    {
        public PPDrawingGroup(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 1036)]
    public class PPDrawing : RegularContainer
    {
        public PPDrawing(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 4008)]
    public class TextBytesAtom : Record
    {
        public static Encoding ENCODING = Encoding.GetEncoding("iso-8859-1");
        public string Text;

        public TextBytesAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            byte[] bytes = new byte[size];
            this.Reader.Read(bytes, 0, (int)size);

            this.Text = new String(ENCODING.GetChars(bytes));
        }

        public override string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Text = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1), this.Text);
        }
    }

    [OfficeRecord(TypeCode = 4080)]
    public class SlideListWithText : RegularContainer
    {
        public SlideListWithText(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    #endregion

    #region Drawing records

    [OfficeRecord(TypeCode = 0xF000)]
    public class DrawingGroup : RegularContainer
    {
        public DrawingGroup(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 0xF002)]
    public class DrawingContainer : RegularContainer
    {
        public DrawingContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 0xF003)]
    public class GroupContainer : RegularContainer
    {
        public GroupContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 0xF004)]
    public class ShapeContainer : RegularContainer
    {
        public ShapeContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 0xF006)]
    public class DrawingGroupRecord : Record
    {
        public class FileIdCluster
        {
            public UInt32 DrawingGroupId;
            public UInt32 CSpIdCur;

            public FileIdCluster(BinaryReader reader)
            {
                this.DrawingGroupId = reader.ReadUInt32();
                this.CSpIdCur = reader.ReadUInt32();
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

        public DrawingGroupRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.MaxShapeId = this.Reader.ReadUInt32();
            this.IdClustersCount = this.Reader.ReadUInt32() - 1; // Office saves the actual value + 1 -- flgr
            this.ShapesSavedCount = this.Reader.ReadUInt32();
            this.DrawingsSavedCount = this.Reader.ReadUInt32();

            for (int i = 0; i < this.IdClustersCount; i++)
            {
                Clusters.Add(new FileIdCluster(this.Reader));
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

    [OfficeRecord(TypeCode = 0xF00D)]
    public class ClientTextbox : RegularContainer
    {
        public ClientTextbox(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(TypeCode = 0xF011)]
    public class ClientData : RegularContainer
    {
        public ClientData(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    #endregion
}
