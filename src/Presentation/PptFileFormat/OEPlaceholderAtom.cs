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
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public enum PlaceholderId
    {
        None = 0,
        MasterTitle = 1,
        MasterBody = 2,
        MasterCenteredTitle = 3,
        MasterSubtitle = 4,
        MasterNotesSlideImage = 5,
        MasterNotesBody = 6,
        MasterDate = 7,
        MasterSlideNumber = 8,
        MasterFooter = 9,
        MasterHeader = 10,
        NotesSlideImage = 11,
        NotesBody = 12,
        Title = 13,
        Body = 14,
        CenteredTitle = 15,
        Subtitle = 16,
        VerticalTextTitle = 17,
        VerticalTextBody = 18,
        Object = 19, // no matter the size
        Graph = 20,
        Table = 21,
        ClipArt = 22,
        OrganizationChart = 23,
        MediaClip = 24
    };

    [OfficeRecordAttribute(3011)]
    public class OEPlaceHolderAtom : Record
    {
        /// <summary>
        /// The placement Id is a number assigned to the placeholder. It goes from -1 to the number of placeholders. See note below.
        /// </summary>
        public Int32 PlacementId; 

        /// <summary>
        /// Type of placeholder. See the Placeholder ID Values table below for valid values.
        /// </summary>
        public PlaceholderId PlaceholderId;

        /// <summary>
        /// Size of the placeholder, which can be:
        ///     0 - full size
        ///     1 - half size
        ///     2 - quart of the slide
        /// </summary>
        public byte PlaceholderSize;

        public OEPlaceHolderAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.PlacementId = this.Reader.ReadInt32();
            this.PlaceholderId = (PlaceholderId) this.Reader.ReadByte();
            this.PlaceholderSize = this.Reader.ReadByte();
            // Throw away additional junk
            this.Reader.ReadUInt16();
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}PlacementId = {2}\n{1}PlaceholderId = {3}, PlaceholderSize = {4})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.PlacementId, this.PlaceholderId, this.PlaceholderSize);
        }

        public bool IsObjectPlaceholder()
        {
            return this.PlaceholderId == PlaceholderId.Object;
        }
    }

}
