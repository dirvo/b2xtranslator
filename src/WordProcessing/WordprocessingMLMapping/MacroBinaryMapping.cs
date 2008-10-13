using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Writer;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class MacroBinaryMapping : DocumentMapping
    {
        public MacroBinaryMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.VbaProjectPart)
        {
            _ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            //get the Class IDs of the directories
            Guid macroClsid = new Guid();
            Guid vbaClsid = new Guid();
            foreach (DirectoryEntry entry in doc.Storage.AllEntries)
            {
                if (entry.Path == "\\Macros")
                {
                    macroClsid = entry.ClsId;
                }
                else if(entry.Path == "\\Macros\\VBA")
                {
                    vbaClsid = entry.ClsId;
                }
            }

            //create a new storage
            StructuredStorageWriter storage = new StructuredStorageWriter();
            storage.RootDirectoryEntry.setClsId(macroClsid);

            //copy the VBA directory
            StorageDirectoryEntry vba = storage.RootDirectoryEntry.AddStorageDirectoryEntry("VBA");
            vba.setClsId(vbaClsid);
            foreach (DirectoryEntry entry in doc.Storage.AllStreamEntries)
            {
                if (entry.Path.StartsWith("\\Macros\\VBA"))
                {
                    vba.AddStreamDirectoryEntry(entry.Name, doc.Storage.GetStream(entry.Path));
                }
            }

            //copy the project streams
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECT", doc.Storage.GetStream("\\Macros\\PROJECT"));
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECTwm", doc.Storage.GetStream("\\Macros\\PROJECTwm"));

           //write the storage to the xml part
            storage.write(_targetPart.GetStream());
        }
    }
}
