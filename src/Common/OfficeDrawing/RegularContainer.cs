using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// Regular containers are containers with Record children.<br/>
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

                child.VerifyReadToEnd();

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

}
