using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using System.Reflection;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    public abstract class OfficeGraphBiffRecord
    {
        RecordNumber _id;
        UInt32 _length;
        long _offset;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public OfficeGraphBiffRecord(IStreamReader reader, RecordNumber id, UInt32 length)
        {
            _reader = reader;
            _offset = _reader.BaseStream.Position;

            _id = id;
            _length = length;
        }

        private static Dictionary<UInt16, Type> TypeToRecordClassMapping = new Dictionary<UInt16, Type>();

        static OfficeGraphBiffRecord()
        {
            UpdateTypeToRecordClassMapping(
                Assembly.GetExecutingAssembly(), 
                typeof(OfficeGraphBiffRecord).Namespace);
        }

        public static void UpdateTypeToRecordClassMapping(Assembly assembly, String ns)
        {
            foreach (Type t in assembly.GetTypes())
            {
                if (ns == null || t.Namespace == ns)
                {
                    object[] attrs = t.GetCustomAttributes(typeof(OfficeGraphBiffRecordAttribute), false);

                    OfficeGraphBiffRecordAttribute attr = null;

                    if (attrs.Length > 0)
                        attr = attrs[0] as OfficeGraphBiffRecordAttribute;

                    if (attr != null)
                    {
                        // Add the type codes of the array
                        foreach (UInt16 typeCode in attr.TypeCodes)
                        {
                            if (TypeToRecordClassMapping.ContainsKey(typeCode))
                            {
                                throw new Exception(String.Format(
                                    "Tried to register TypeCode {0} to {1}, but it is already registered to {2}",
                                    typeCode, t, TypeToRecordClassMapping[typeCode]));
                            }
                            TypeToRecordClassMapping.Add(typeCode, t);
                        }
                    }
                }
            }
        }

        public static OfficeGraphBiffRecord ReadRecord(IStreamReader reader)
        {
            OfficeGraphBiffRecord result = null;
            try
            {
                UInt16 id = reader.ReadUInt16();
                UInt16 size = reader.ReadUInt16();
                Type cls;
                if (TypeToRecordClassMapping.TryGetValue(id, out cls))
                {
                    ConstructorInfo constructor = cls.GetConstructor(
                        new Type[] { typeof(IStreamReader), typeof(RecordNumber), typeof(UInt16) }
                        );

                    try
                    {
                        result = (OfficeGraphBiffRecord)constructor.Invoke(
                            new object[] {reader, id, size }
                            );
                    }
                    catch (TargetInvocationException e)
                    {
                        throw e.InnerException;
                    }
                }
                else
                {
                    result = new UnknownGraphRecord(reader, id, size);
                }

                return result;
            }
            catch (OutOfMemoryException e)
            {
                throw new Exception("Invalid record", e);
            }
        }

        public RecordNumber Id
        {
            get { return _id; }
        }

        public UInt32 Length
        {
            get { return _length; }
        }

        public long Offset
        {
            get { return _offset; }
        }

        IStreamReader _reader;
        public IStreamReader Reader
        {
            get { return _reader; }
            set { this._reader = value; }
        }
    }
}
