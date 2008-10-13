/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class OleObject : IVisitable
    {
        public enum LinkUpdateOption
        {
            NoLink = 0,
            Always = 1,
            OnCall = 3
        }

        public string ObjectId;

        public Guid ClassId;

        /// <summary>
        /// The path of the object in the storage
        /// </summary>
        public string Path;

        /// <summary>
        /// The the value is true, the object is a linked object
        /// </summary>
        public bool fLinked;

        /// <summary>
        /// Display name of the linked object or embedded object.
        /// </summary>
        public string UserType;

        public string ClipboardFormat;

        public string Link;

        public string Program;

        public LinkUpdateOption UpdateMode;

        public Dictionary<string, VirtualStream> Streams;

        private StructuredStorageReader _docStorage;

        public OleObject(CharacterPropertyExceptions chpx, StructuredStorageReader docStorage)
        {
            this._docStorage = docStorage;
            this.ObjectId = getOleEntryName(chpx);

            this.Path = "\\ObjectPool\\" + this.ObjectId + "\\";
            processOleStream(this.Path + "\u0001Ole");

            if (this.fLinked)
            {
                processLinkInfoStream(this.Path + "\u0003LinkInfo");
            }
            else
            {
                processCompObjStream(this.Path + "\u0001CompObj");
            }

            //get the storage entries of this object
            this.Streams = new Dictionary<string, VirtualStream>();
            foreach (string streamname in docStorage.FullNameOfAllStreamEntries)
            {
                if (streamname.StartsWith(this.Path))
                {
                    this.Streams.Add(streamname.Substring(streamname.LastIndexOf("\\") + 1), docStorage.GetStream(streamname));
                }
            }

            //find the class if of this object
            foreach (DirectoryEntry entry in docStorage.AllEntries)
            {
                if (entry.Name == this.ObjectId)
                {
                    this.ClassId = entry.ClsId;
                    break;
                }
            }
        }

        private void processLinkInfoStream(string linkStream)
        {
            try
            {
                VirtualStreamReader reader = new VirtualStreamReader(_docStorage.GetStream(linkStream));

                //there are two versions of the Link string, one contains ANSI characters, the other contains
                //unicode characters.
                //Both strings seem not to be standardized:
                //The length prefix is a character count EXCLUDING the terminating zero

                //Read the ANSI version
                Int16 cch = reader.ReadInt16();
                byte[] str = reader.ReadBytes(cch);
                this.Link = Encoding.ASCII.GetString(str);
                
                //skip the terminating zero of the ANSI string
                //even if the characters are ANSI chars, the terminating zero has 2 bytes
                reader.ReadBytes(2);

                //skip the next 4 bytes (flags?)
                reader.ReadBytes(4);

                //Read the Unicode version
                cch = reader.ReadInt16();
                str = reader.ReadBytes(cch * 2);
                this.Link = Encoding.Unicode.GetString(str);

                //skip the terminating zero of the Unicode string
                reader.ReadBytes(2);
            }
            catch (StreamNotFoundException) { }
        }

        private void processCompObjStream(string compStream)
        {
            try
            {
                VirtualStreamReader reader = new VirtualStreamReader(_docStorage.GetStream(compStream));

                //skip the CompObjHeader
                reader.ReadBytes(28);

                this.UserType = Utils.ReadLengthPrefixedAnsiString(reader.BaseStream);
                this.ClipboardFormat = Utils.ReadLengthPrefixedAnsiString(reader.BaseStream);
                this.Program = Utils.ReadLengthPrefixedAnsiString(reader.BaseStream);
            }
            catch (StreamNotFoundException) { }
        }

        private void processOleStream(string oleStream)
        {
            try
            {
                VirtualStreamReader reader = new VirtualStreamReader(_docStorage.GetStream(oleStream));

                //skip version
                reader.ReadBytes(4);

                //read the embedded/linked flag
                Int32 flag = reader.ReadInt32();
                this.fLinked = Utils.BitmaskToBool(flag, 0x1);

                //Link update option
                this.UpdateMode = (LinkUpdateOption)reader.ReadInt32();
            }
            catch (StreamNotFoundException) { }
        }

        private string getOleEntryName(CharacterPropertyExceptions chpx)
        {
            string ret = null;

            foreach (SinglePropertyModifier sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCPicLocation)
                {
                    ret = "_" + System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    break;
                }
            }

            return ret;
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<OleObject>)mapping).Apply(this);
        }

        #endregion
    }
}
