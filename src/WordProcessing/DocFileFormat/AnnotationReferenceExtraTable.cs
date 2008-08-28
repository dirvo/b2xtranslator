using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AnnotationReferenceExtraTable : List<AnnotationReferenceDescriptorExtra>
    {
        public AnnotationReferenceExtraTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            tableStream.Seek((long)fib.fcAtrdExtra, System.IO.SeekOrigin.Begin);
            VirtualStreamReader reader = new VirtualStreamReader(tableStream);

            int cbATRDPost10 = 16;
            int n = (int)fib.lcbAtrdExtra / cbATRDPost10;

            //read the n ATRDPost10 structs
            for (int i = 0; i < n; i++)
            {
                this.Add(new AnnotationReferenceDescriptorExtra(reader));        
            }
        }
    }
}
