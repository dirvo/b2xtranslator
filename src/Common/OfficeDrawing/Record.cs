using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;
using System.Reflection;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class Record : IEnumerable<Record>
    {
        public const uint HEADER_SIZE_IN_BYTES = (16 + 16 + 32) / 8;

        public uint TotalSize
        {
            get { return HeaderSize + BodySize; }
        }

        private Record _ParentRecord = null;

        public Record ParentRecord
        {
            get { return _ParentRecord; }
            set
            {
                if (_ParentRecord != null)
                    throw new Exception("Can only set ParentRecord once");

                _ParentRecord = value;
                this.AfterParentSet();
            }
        }

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

        public virtual void AfterParentSet() { }

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

        public void VerifyReadToEnd()
        {
            long streamPos = this.Reader.BaseStream.Position;
            long streamLen = this.Reader.BaseStream.Length;

            if (streamPos != streamLen)
            {
                throw new Exception(String.Format(
                    "Record {3} didn't read to end: (stream position: {1} of {2})\n{0}",
                    this, streamPos, streamLen, this.GetIdentifier()));
            }
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

        protected static string IndentationForDepth(uint depth)
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
                if (t.Namespace == typeof(Record).Namespace)
                {
                    object[] attrs = t.GetCustomAttributes(typeof(OfficeRecordAttribute), false);

                    OfficeRecordAttribute attr = null;

                    if (attrs.Length > 0)
                        attr = attrs[0] as OfficeRecordAttribute;

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

}
