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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Reflection;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class PowerpointDocument : BinaryDocument, IVisitable, IEnumerable<Record>
    {
        static PowerpointDocument() {
            Record.UpdateTypeToRecordClassMapping(Assembly.GetExecutingAssembly(), typeof(PowerpointDocument).Namespace);
        }

        /// <summary>
        /// The stream "PowerPoint Document"
        /// </summary>
        public VirtualStream PowerpointDocumentStream;

        /// <summary>
        /// The stream "Current User"
        /// </summary>
        public VirtualStream CurrentUserStream;

        /// <summary>
        /// Atom containing information about the last user to edit this document and a reference to that last edit.
        /// </summary>
        public CurrentUserAtom CurrentUserAtom;

        /// <summary>
        /// The stream "Pictures"
        /// </summary>
        public VirtualStream PicturesStream;

        /// <summary>
        /// Container containing all media elements inside the Pictures stream.
        /// </summary>
        public Pictures PicturesContainer;

        /// <summary>
        /// The last edit done to this document.
        /// </summary>
        public UserEditAtom LastUserEdit;

        /// <summary>
        /// The persist object directory is used for mapping persist object identifiers to document stream offsets.
        /// </summary>
        public Dictionary<UInt32, UInt32> PersistObjectDirectory = new Dictionary<uint,uint>();

        /// <summary>
        /// The DocumentContainer record for this document.
        /// </summary>
        public DocumentContainer DocumentRecord;

        /// <summary>
        /// List of all main (regular) master records for this document.
        /// </summary>
        public List<MainMaster> MainMasterRecords = new List<MainMaster>();

        /// <summary>
        /// List of title master records for this document.
        /// </summary>
        public List<Slide> TitleMasterRecords = new List<Slide>();

        /// <summary>
        /// Dictionary used for finding MasterRecords (title / main masters) by master id.
        /// </summary>
        private Dictionary<UInt32, Slide> MasterRecordsById =
            new Dictionary<UInt32, Slide>();

        /// <summary>
        /// List of all slide records for this document.
        /// </summary>
        public List<Slide> SlideRecords = new List<Slide>();

        public PowerpointDocument(StructuredStorageReader file)
        {
            try
            {
                this.CurrentUserStream = file.GetStream("Current User");
                this.CurrentUserAtom = (CurrentUserAtom)Record.ReadRecord(this.CurrentUserStream, 0);
            }
            catch (InvalidRecordException e)
            {
                throw new InvalidStreamException("Current user stream is not valid", e);
            }

            // Optional 'Pictures' stream
            if (file.FullNameOfAllStreamEntries.Contains("\\Pictures"))
            {
                try
                {
                    this.PicturesStream = file.GetStream("Pictures");
                    this.PicturesContainer = new Pictures(new BinaryReader(this.PicturesStream), (uint)this.PicturesStream.Length, 0, 0, 0);
                }
                catch (InvalidRecordException e)
                {
                    throw new InvalidStreamException("Pictures stream is not valid", e);
                }
            }

            this.PowerpointDocumentStream = file.GetStream("PowerPoint Document");
            this.PowerpointDocumentStream.Seek(this.CurrentUserAtom.OffsetToCurrentEdit, SeekOrigin.Begin);

            this.LastUserEdit = (UserEditAtom)Record.ReadRecord(this.PowerpointDocumentStream, 0);

            this.ConstructPersistObjectDirectory();

            this.IdentifyDocumentPersistObject();
            this.IdentifyMasterPersistObjects();
            this.IdentifySlidePersistObjects();

            // TODO: notes / handout masters and slides
        }

        /// <summary>
        /// Returns the slide or main master with the specified masterId or null if none exists.
        /// </summary>
        /// <param name="masterId">id of master to find</param>
        /// <returns>Slide or main master with the specified masterId or null if none exists</returns>
        public Slide FindMasterRecordById(UInt32 masterId)
        {
            return this.MasterRecordsById[masterId];
        }

        /// <summary>
        /// Tries to find a record with the supplied persistId and type in the PersistObjectDirectory, reads it and returns it.
        /// </summary>
        /// <typeparam name="T">Type of record</typeparam>
        /// <param name="persistId">persist id of record to look up</param>
        /// <returns>Matching record of given type or null</returns>
        public T GetPersistObject<T>(uint persistId) where T : Record
        {
            if (!this.PersistObjectDirectory.ContainsKey(persistId))
                return null;

            UInt32 offset = this.PersistObjectDirectory[persistId];
            this.PowerpointDocumentStream.Seek(offset, SeekOrigin.Begin);
            return (T)Record.ReadRecord(this.PowerpointDocumentStream, 0);
        }

        /// <summary>
        /// Find the root DocumentContainer record for this presentation.
        /// 
        /// This is done by looking up the document persist id reference of the last user edit in the persist object directory.
        /// </summary>
        private void IdentifyDocumentPersistObject()
        {
            this.DocumentRecord = this.GetPersistObject<DocumentContainer>(this.LastUserEdit.DocPersistIdRef);
        }

        /// <summary>
        /// Find all master records for this presentation.
        /// 
        /// This is done by looking up all persist id references of all SlidePersistAtoms of the DocumentRecord's MasterPersistList
        /// in the persist object directory.
        /// </summary>
        private void IdentifyMasterPersistObjects()
        {
            foreach (SlidePersistAtom masterPersistAtom in this.DocumentRecord.MasterPersistList)
            {
                Slide master = this.GetPersistObject<Slide>(masterPersistAtom.PersistIdRef);
                master.PersistAtom = masterPersistAtom;

                if (master is MainMaster)
                    this.MainMasterRecords.Add((MainMaster)master);
                else
                    this.TitleMasterRecords.Add(master);

                this.MasterRecordsById.Add(master.PersistAtom.SlideId, master);
            }
        }

        /// <summary>
        /// Find all Slide records for this presentation.
        /// 
        /// This is done by looking up all persist id references of all SlidePersistAtoms of the DocumentRecord's MasterPersistList
        /// in the persist object directory.
        /// </summary>
        private void IdentifySlidePersistObjects()
        {
            foreach (SlidePersistAtom slidePersistAtom in this.DocumentRecord.SlidePersistList)
            {
                Slide slide = this.GetPersistObject<Slide>(slidePersistAtom.PersistIdRef);
                slide.PersistAtom = slidePersistAtom;
                this.SlideRecords.Add(slide);
            }
        }

        /// <summary>
        /// Construct the complete persist object directory by traversing all PersistDirectoryAtoms
        /// from all UserEditAtoms from the last edit to the first one and adding all _entries of
        /// all encountered persist directories to the resulting persist object directory.
        /// 
        /// When the same persist id occurs multiple times with different offsets the one from the
        /// last user edit will end up in the persist object directory, overwriting earlier edits.
        /// </summary>
        private void ConstructPersistObjectDirectory()
        {
            List<PersistDirectoryAtom> pdAtoms = FindLivePersistDirectoryAtoms();

            foreach (PersistDirectoryAtom pdAtom in pdAtoms)
            {
                foreach (PersistDirectoryEntry pdEntry in pdAtom.PersistDirEntries)
                {
                    uint pid = pdEntry.StartPersistId;

                    foreach (UInt32 poff in pdEntry.PersistOffsetEntries)
                    {
                        this.PersistObjectDirectory[pid] = poff;
                        pid++;
                    }
                }
            }
        }

        /// <summary>
        /// Find all live PersistDirectoryAtoms by traversing all UserEditAtoms starting from CurrentUserAtom.
        /// </summary>
        /// <returns>List of PersistDirectoryAtoms. The oldest PersitDirectoryAtom will be the first element of the list.</returns>
        private List<PersistDirectoryAtom> FindLivePersistDirectoryAtoms()
        {
            List<PersistDirectoryAtom> result = new List<PersistDirectoryAtom>();

            UserEditAtom userEditAtom = this.LastUserEdit;

            while (userEditAtom != null)
            {
                this.PowerpointDocumentStream.Seek(userEditAtom.OffsetPersistDirectory, SeekOrigin.Begin);
                PersistDirectoryAtom pdAtom = (PersistDirectoryAtom)Record.ReadRecord(this.PowerpointDocumentStream, 0);
                result.Insert(0, pdAtom);

                this.PowerpointDocumentStream.Seek(userEditAtom.OffsetLastEdit, SeekOrigin.Begin);

                if (userEditAtom.OffsetLastEdit != 0)
                    userEditAtom = (UserEditAtom)Record.ReadRecord(this.PowerpointDocumentStream, 0);
                else
                    userEditAtom = null;
            }

            return result;
        }

        override public string ToString()
        {
            StringBuilder result = new StringBuilder(base.ToString());

            result.Append("CurrentUserAtom: ");
            result.AppendLine(this.CurrentUserAtom.ToString());
            result.AppendLine();

            result.Append("DocumentRecord: ");
            result.AppendLine(this.DocumentRecord.ToString());

            foreach (Record r in this.MainMasterRecords)
            {
                result.AppendLine();
                result.Append("MainMasterRecord: ");
                result.AppendLine(r.ToString());
            }

            foreach (Record r in this.TitleMasterRecords)
            {
                result.AppendLine();
                result.Append("TitleMasterRecord: ");
                result.AppendLine(r.ToString());
            }

            foreach (Record r in this.SlideRecords)
            {
                result.AppendLine();
                result.Append("SlideRecord: ");
                result.AppendLine(r.ToString());
            }

            return result.ToString();
        }

        #region IVisitable Members

        override public void Convert<T>(T mapping)
        {
            ((IMapping<PowerpointDocument>)mapping).Apply(this);
        }

        #endregion

        #region IEnumerable<Record> Member

        public IEnumerator<Record> GetEnumerator()
        {
            foreach (UInt32 persistId in this.PersistObjectDirectory.Keys)
            {
                yield return this.GetPersistObject<Record>(persistId);
            }
        }

        #endregion

        #region IEnumerable Member

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            IEnumerator<Record> e = this.GetEnumerator();
            return (System.Collections.IEnumerator)e;
        }

        #endregion
    }
}
