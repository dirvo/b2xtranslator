using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.Reflection;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class Plex
    {
        protected const int CP_LENGTH = 4;

        public List<Int32> CharacterPositions;
        public List<PlexStruct> Structs;

        public Plex(Type structType, int structureLength, VirtualStream tableStream, Int32 fc, UInt32 lcb)
        {
            tableStream.Seek((long)fc, System.IO.SeekOrigin.Begin);
            VirtualStreamReader reader = new VirtualStreamReader(tableStream);

            int n = ((int)lcb - CP_LENGTH) / (structureLength + CP_LENGTH);

            //read the n + 1 CPs
            this.CharacterPositions = new List<Int32>();
            for (int i = 0; i < n + 1; i++)
            {
                this.CharacterPositions.Add(reader.ReadInt32());
            }

            //read the n structs
            this.Structs = new List<PlexStruct>(); 
            for (int i = 0; i < n; i++)
            {
                ConstructorInfo constructor = structType.GetConstructor(new Type[] { typeof(VirtualStreamReader) });
                PlexStruct st = (PlexStruct)constructor.Invoke(new object[] { reader });
                this.Structs.Add(st);
            }
        }

        /// <summary>
        /// Retruns the struct that matches the given character position.
        /// </summary>
        /// <param name="cp">The character position</param>
        /// <returns>The matching struct</returns>
        public PlexStruct GetStruct(Int32 cp)
        {
            int index = -1;
            for (int i = 0; i < this.CharacterPositions.Count; i++)
            {
                if (this.CharacterPositions[i] == cp)
                {
                    index = i;
                    break;
                }
            }

            if (index >= 0 && index < this.Structs.Count)
            {
                return this.Structs[index];
            }
            else
            {
                return null;
            }
        }
    }
}
