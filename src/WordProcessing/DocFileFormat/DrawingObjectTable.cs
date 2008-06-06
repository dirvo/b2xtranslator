using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class DrawingObjectTable
    {
        public DrawingGroup drawingGroup;
        public DrawingContainer drawing;

        public DrawingObjectTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            VirtualStreamReader reader = new VirtualStreamReader(tableStream);
            tableStream.Seek(fib.fcDggInfo, System.IO.SeekOrigin.Begin);

            this.drawingGroup = (DrawingGroup)Record.readRecord(reader);

            //word always writes an empty byte between the two record ...
            //I don't know why ...
            reader.ReadByte(); 

            this.drawing = (DrawingContainer)Record.readRecord(reader);
        }

        /// <summary>
        /// Searches the matching shape
        /// </summary>
        /// <param name="spid">The shape ID</param>
        /// <returns>The ShapeContainer</returns>
        public ShapeContainer GetShapeContainer(int spid)
        {
            ShapeContainer ret = null;

            //get the shape from the DrawingObjectTable
            GroupContainer group = (GroupContainer)this.drawing.Children[1];
            for(int i=1; i<group.Children.Count; i++)
            {
                Record groupChild = group.Children[i];
                if(groupChild.TypeCode == 0xF003)
                {
                    //it's a group of shapes
                }
                else if (groupChild.TypeCode == 0xF004)
                {
                    ShapeContainer container = (ShapeContainer)groupChild;
                    Shape shape = (Shape)container.Children[0];
                    if (shape.spid == spid)
                    {
                        ret = container;
                    }
                }
            }

            return ret;
        }
    }
}
