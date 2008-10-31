using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AnnotationOwnerList : List<string>
    {
        public AnnotationOwnerList(FileInformationBlock fib, VirtualStream tableStream) : base()
        {
            tableStream.Seek(fib.fcGrpXstAtnOwners, System.IO.SeekOrigin.Begin);

            while (tableStream.Position < (fib.fcGrpXstAtnOwners + fib.lcbGrpXstAtnOwners))
            {
                this.Add(Utils.ReadXstz(tableStream));
            }
        }
    }
}
