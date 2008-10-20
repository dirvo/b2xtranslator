using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class CustomToolbarWrapper : ByteStructure
    {
        /// <summary>
        /// Signed integer that specifies the size in bytes of a TBDelta structure. <br/>
        /// MUST be 0x0012.
        /// </summary>
        public Int16 cbTBD;

        /// <summary>
        /// Signed integer that specifies the number of elements in the rCustomizations array. <br/>
        /// MUST be greater than 0x0000.
        /// </summary>
        public Int16 cCust;

        /// <summary>
        /// Signed integer that specifies the size, in bytes, of the rtbdc array.<br/> 
        /// MUST be greater or equal to 0x00000000.
        /// </summary>
        public Int32 cbDTBC;

        /// <summary>
        /// 
        /// </summary>
        public List<ToolbarControl> rtbdc;

        /// <summary>
        /// 
        /// </summary>
        public List<ToolbarCustomization> rCustomizations;

        public CustomToolbarWrapper(VirtualStreamReader reader) : base(reader, ByteStructure.VARIABLE_LENGTH)
        {
            long startPos = reader.BaseStream.Position;

            //skip the first 7 bytes
            byte[] skipped = reader.ReadBytes(7);

            this.cbTBD = reader.ReadInt16();
            this.cCust = reader.ReadInt16();
            this.cbDTBC = reader.ReadInt32();

            this.rtbdc = new List<ToolbarControl>();
            int max = (int)(reader.BaseStream.Position + cbDTBC);
            while (reader.BaseStream.Position < max)
            {
                this.rtbdc.Add(new ToolbarControl(reader));
            }

            this.rCustomizations = new List<ToolbarCustomization>();
            for (int i = 0; i < cCust; i++)
			{
			  this.rCustomizations.Add(new ToolbarCustomization(reader));
			}

            long endPos = reader.BaseStream.Position;

            //read the raw bytes
            reader.BaseStream.Seek(startPos - 1, System.IO.SeekOrigin.Begin);
            this._rawBytes = reader.ReadBytes((int)(endPos - startPos + 1));
        }
    }
}
