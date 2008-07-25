/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.Spreadsheet
{
    public class WorkbookPart : OpenXmlPart
    {
        private UInt16 WorksheetNumber;
        private UInt16 ExternalLinkNumber;
        protected WorksheetPart workSheetPart;
        protected SharedStringPart sharedStringPart;
        protected ExternalLinkPart externalLinkPart; 

        public WorkbookPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
            this.WorksheetNumber = 1;
            this.ExternalLinkNumber = 1; 
        }

        public override string ContentType
        {
            get { return SpreadsheetMLContentTypes.Workbook; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.OfficeDocument; }
        }

        /// <summary>
        /// returns the worksheet part from the new excel document 
        /// </summary>
        /// <returns></returns>
        public WorksheetPart AddWorksheetPart()
        {
            this.workSheetPart = new WorksheetPart(this, this.WorksheetNumber);
            this.WorksheetNumber++;
            return this.AddPart(this.workSheetPart);
        }

        /// <summary>
        /// return the latest created worksheetpart
        /// </summary>
        /// <returns></returns>
        public WorksheetPart GetWorksheetPart()
        {
            return this.workSheetPart; 
        }

        /// <summary>
        /// returns the worksheet part from the new excel document 
        /// </summary>
        /// <returns></returns>
        public ExternalLinkPart AddExternalLinkPart()
        {
            this.externalLinkPart = new ExternalLinkPart(this, this.ExternalLinkNumber);
            this.ExternalLinkNumber++;
            return this.AddPart(this.externalLinkPart);
        }

        /// <summary>
        /// return the latest created worksheetpart
        /// </summary>
        /// <returns></returns>
        public ExternalLinkPart GetExternalLinkPart()
        {
            return this.externalLinkPart;
        }

        public override string TargetName { get { return "workbook"; } }
        public override string TargetDirectory { get { return "xl"; } }


        /// <summary>
        /// returns the sharedstringtable part from the new excel document 
        /// </summary>
        /// <returns></returns>
        public SharedStringPart AddSharedStringPart()
        {
            this.sharedStringPart = new SharedStringPart(this);
            return this.AddPart(this.sharedStringPart);
        }
    }
}

