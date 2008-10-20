using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class KeyMapEntry : ByteStructure
    {
        public enum ActionType
        {
            ktCid,
            ktChar,
            ktMask
        }

        
        /// <summary>
        /// 
        /// </summary>
        public Int16 kcm1;

        /// <summary>
        /// 
        /// </summary>
        public Int16 kcm2;

        /// <summary>
        /// 
        /// </summary>
        public ActionType kt;

        /// <summary>
        /// 
        /// </summary>
        public char paramChar;

        public CommandIdentifier paramCid;

        private const int KME_LENGTH = 14;
        
        public KeyMapEntry(VirtualStreamReader reader)
            : base(reader, KME_LENGTH)
        {
            //ignore the first 4 bytes
            reader.ReadBytes(4);

            //Primary KCM
            this.kcm1 = reader.ReadInt16();

            //Secondary KCM
            this.kcm2 = reader.ReadInt16();

            //Key Action Type
            this.kt = (ActionType)reader.ReadInt16();

            //read the params
            switch (kt)
            {
                case ActionType.ktCid:
                    this.paramCid = new CommandIdentifier(reader);
                    break;
                case ActionType.ktChar:
                    this.paramChar = (char)reader.ReadInt32();
                    break;
                default:
                    reader.ReadBytes(4);
                    break;
            }
        }
    }
}
