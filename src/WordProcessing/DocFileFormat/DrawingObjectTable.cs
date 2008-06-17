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
        public List<DrawingContainer> drawings;

        public DrawingObjectTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            VirtualStreamReader reader = new VirtualStreamReader(tableStream);
            tableStream.Seek(fib.fcDggInfo, System.IO.SeekOrigin.Begin);

            if (fib.lcbDggInfo > 0)
            {
                int maxPosition = (int)(fib.fcDggInfo + fib.lcbDggInfo);

                Record dggRec = Record.readRecord(reader);
                this.drawingGroup = (DrawingGroup)dggRec;
                this.drawings = new List<DrawingContainer>();

                while (reader.BaseStream.Position < maxPosition)
                {
                    //word writes an empty byte between the two record ...
                    //I don't know why ...
                    reader.ReadByte();
                    this.drawings.Add((DrawingContainer)Record.readRecord(reader));
                }
            }
        }

        /// <summary>
        /// Searches the matching shape
        /// </summary>
        /// <param name="spid">The shape ID</param>
        /// <returns>The ShapeContainer</returns>
        public ShapeContainer GetShapeContainer(int spid)
        {
            ShapeContainer ret = null;

            foreach (DrawingContainer drawing in this.drawings)
            {
                //get the shape from the DrawingObjectTable
                //GroupContainer group = (GroupContainer)drawing.Children[1];
                GroupContainer group = (GroupContainer)drawing.FirstChildWithType<GroupContainer>();

                if (group != null)
                {
                    for (int i = 1; i < group.Children.Count; i++)
                    {
                        Record groupChild = group.Children[i];
                        if (groupChild.TypeCode == 0xF003)
                        {
                            //It's a group of shapes
                            GroupContainer subgroup = (GroupContainer)groupChild;

                            //the referenced shape must be the first shape in the group
                            ShapeContainer container = (ShapeContainer)subgroup.Children[0];
                            Shape shape = (Shape)container.Children[1];
                            if (shape.spid == spid)
                            {
                                ret = container;
                                break;
                            }
                        }
                        else if (groupChild.TypeCode == 0xF004)
                        {
                            //It's a singe shape
                            ShapeContainer container = (ShapeContainer)groupChild;
                            Shape shape = (Shape)container.Children[0];
                            if (shape.spid == spid)
                            {
                                ret = container;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    continue;
                }

                if (ret != null)
                {
                    break;
                }
            }

            return ret;
        }
    }
}
