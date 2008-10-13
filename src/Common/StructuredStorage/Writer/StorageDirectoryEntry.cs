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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    public class StorageDirectoryEntry : BaseDirectoryEntry
    {
        List<StreamDirectoryEntry> _streamDirectoryEntries = new List<StreamDirectoryEntry>();
        internal List<StreamDirectoryEntry> StreamDirectoryEntries
        {
            get { return _streamDirectoryEntries; }            
        }


        List<StorageDirectoryEntry> _storageDirectoryEntries = new List<StorageDirectoryEntry>();
        internal List<StorageDirectoryEntry> StorageDirectoryEntries
        {
            get { return _storageDirectoryEntries; }            
        }

        List<BaseDirectoryEntry> _allDirectoryEntries = new List<BaseDirectoryEntry>();


        internal StorageDirectoryEntry(string name, StructuredStorageContext context)
            : base(name, context)
        {
            Type = DirectoryEntryType.STGTY_STORAGE;
        }


        public void AddStreamDirectoryEntry(string name, Stream stream)
        {
            if (_streamDirectoryEntries.Exists(delegate(StreamDirectoryEntry a) { return name == a.Name; }))
            {
                return;
            }
            StreamDirectoryEntry newDirEntry = new StreamDirectoryEntry(name, stream, Context);
            _streamDirectoryEntries.Add(newDirEntry);
            _allDirectoryEntries.Add(newDirEntry);
        }


        public StorageDirectoryEntry AddStorageDirectoryEntry(string name)
        {
            StorageDirectoryEntry result = null;
            result = _storageDirectoryEntries.Find(delegate(StorageDirectoryEntry a) { return name == a.Name; });
            if (result != null)
            {
                // entry exists
                return result;
            }
            result = new StorageDirectoryEntry(name, Context);
            _storageDirectoryEntries.Add(result);
            _allDirectoryEntries.Add(result);
            return result;
        }


        public void setClsId(Guid clsId)
        {
            ClsId = clsId;
        }


        internal List<BaseDirectoryEntry> RecursiveGetAllDirectoryEntries()
        {
            List<BaseDirectoryEntry> result = new List<BaseDirectoryEntry>();
            return RecursiveGetAllDirectoryEntries(result);
        }


        private List<BaseDirectoryEntry> RecursiveGetAllDirectoryEntries(List<BaseDirectoryEntry> result)
        {
            foreach (StorageDirectoryEntry entry in _storageDirectoryEntries)
            {
                result.AddRange(entry.RecursiveGetAllDirectoryEntries());
            }
            foreach (StreamDirectoryEntry entry in _streamDirectoryEntries)
            {
                result.Add(entry);
            }
            if (!result.Contains(this))
            {
                result.Add(this);
            }

            return result;
        }


        internal void RecursiveCreateRedBlackTrees()
        {
            this.ChildSiblingSid = CreateRedBlackTree();

            foreach (StorageDirectoryEntry entry in _storageDirectoryEntries)
            {
                entry.RecursiveCreateRedBlackTrees();
            }

            //foreach (BaseDirectoryEntry entry in _allDirectoryEntries)
            //{
            //    UInt32 left = entry.LeftSiblingSid;
            //    UInt32 right = entry.RightSiblingSid;
            //    UInt32 child = entry.ChildSiblingSid;
            //    Console.WriteLine("{0:X02}: Left: {2:X02}, Right: {3:X02}, Child: {4:X02}, Name: {1}, Color: {5}", entry.Sid, entry.Name, (left > 0xFF) ? 0xFF : left, (right > 0xFF) ? 0xFF : right, (child > 0xFF) ? 0xFF : child, entry.Color.ToString());
            //}
            //Console.WriteLine("----------");
        }


        private UInt32 CreateRedBlackTree()
        {          
            //_allDirectoryEntries.Sort(
            //    delegate(BaseDirectoryEntry a, BaseDirectoryEntry b) 
            //    { return (a.Name.Length == b.Name.Length) ? a.Name.ToLower().CompareTo(b.Name.ToLower()) : a.Name.Length.CompareTo(b.Name.Length); }
            //    );

            _allDirectoryEntries.Sort(DirectoryEntryComparison);

            foreach (BaseDirectoryEntry entry in _allDirectoryEntries)
            {
                entry.Sid = Context.getNewSid();
            }

            //this.ChildSiblingSid = setRelationsAndColorRecursive(this._allDirectoryEntries, (int)Math.Floor(Math.Log(_allDirectoryEntries.Count, 2)), 0);
            return setRelationsAndColorRecursive(this._allDirectoryEntries, (int)Math.Floor(Math.Log(_allDirectoryEntries.Count, 2)), 0);
        }


        private UInt32 setRelationsAndColorRecursive(List<BaseDirectoryEntry> entryList, int treeHeight, int treeLevel)
        {
            if (entryList.Count < 1)
            {                
                return SectorId.FREESECT;
            }

            if (entryList.Count == 1)
            {
                if (treeLevel == treeHeight)
                {
                    entryList[0].Color = DirectoryEntryColor.DE_RED;
                }
                return entryList[0].Sid;
            }

            int middleIndex = getMiddleIndex(entryList);
            List<BaseDirectoryEntry> leftSubTree = entryList.GetRange(0, middleIndex);
            List<BaseDirectoryEntry> rightSubTree = entryList.GetRange(middleIndex + 1, entryList.Count - middleIndex - 1);
            int leftmiddleIndex = getMiddleIndex(leftSubTree);
            int rightmiddleIndex = getMiddleIndex(rightSubTree);
            if (leftSubTree.Count > 0)
            {
                entryList[middleIndex].LeftSiblingSid = leftSubTree[leftmiddleIndex].Sid;
                setRelationsAndColorRecursive(leftSubTree, treeHeight, treeLevel + 1);
            }
            if (rightSubTree.Count > 0)
            {
                entryList[middleIndex].RightSiblingSid = rightSubTree[rightmiddleIndex].Sid;
                setRelationsAndColorRecursive(rightSubTree, treeHeight, treeLevel + 1);
            }            

            return entryList[middleIndex].Sid;
        }


        private static int getMiddleIndex(List<BaseDirectoryEntry> list)
        {
            return (int)Math.Floor((list.Count - 1)/ 2.0);
        }

        protected int DirectoryEntryComparison(BaseDirectoryEntry a, BaseDirectoryEntry b)
        {
            if (a.Name.Length != b.Name.Length)
            {
                return a.Name.Length.CompareTo(b.Name.Length);
            }

            String aU = a.Name.ToUpper();
            String bU = b.Name.ToUpper();

            for (int i = 0; i < aU.Length; i++)
            {
                if ((UInt32)aU[i] != (UInt32)bU[i])
                {
                    return ((UInt32)aU[i]).CompareTo((UInt32)bU[i]);
                }
            }

            return 0;
        }
    }
}

