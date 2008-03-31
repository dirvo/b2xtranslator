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
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class WordDocument : IVisitable
    {
        /// <summary>
        /// 
        /// </summary>
        public PieceTable PieceTable;

        /// <summary>
        /// The stream "WordDocument"
        /// </summary>
        public VirtualStream WordDocumentStream;

        /// <summary>
        /// The stream "0Table" or "1Table"
        /// </summary>
        public VirtualStream TableStream;

        /// <summary>
        /// The stream called "Data"
        /// </summary>
        public VirtualStream DataStream;

        /// <summary>
        /// The file information block of the word document
        /// </summary>
        public FileInformationBlock FIB;

        /// <summary>
        /// The text part of the Word document
        /// </summary>
        public List<char> Text;

        /// <summary>
        /// The macros of the Word document
        /// </summary>
        public List<char> Macros;

        /// <summary>
        /// The headers of the Word document
        /// </summary>
        public List<char> Headers;

        /// <summary>
        /// The textboxes of the Word document
        /// </summary>
        public List<char> Textboxes;

        /// <summary>
        /// The annotations of the Word document
        /// </summary>
        public List<char> Annotations;

        /// <summary>
        /// The endnotes of the Word document
        /// </summary>
        public List<char> Endnotes;

        /// <summary>
        /// The footnotes of the Word document
        /// </summary>
        public List<char> Footnotes;

        /// <summary>
        /// The textboxes in headers of the Word document
        /// </summary>
        public List<char> HeaderTextboxes;

        /// <summary>
        /// The style sheet of the document
        /// </summary>
        public StyleSheet Styles;

        /// <summary>
        /// A list of all font names, used in the doucument
        /// </summary>
        public FontTable FontTable;

        /// <summary>
        /// A dictionary with all ParagraphPropertyExceptions.<br/>
        /// The key is the offset where the PAPX starts.
        /// </summary>
        public Dictionary<Int32, ParagraphPropertyExceptions> AllPapx;

        /// <summary>
        /// A dictionary with all CharacterPropertyExceptions.<br/>
        /// The key is the offset where the CHPX starts.
        /// </summary>
        public Dictionary<Int32, CharacterPropertyExceptions> AllChpx;

        public WordDocument(StorageReader reader)
        {
            this.WordDocumentStream = reader.GetStream("WordDocument");

            //parse FIB
            this.FIB = new FileInformationBlock(this.WordDocumentStream);

            if (this.FIB.nFib < 105)
                throw new InvalidFileException("DocFileFormat doesn't support Word versions older than Word 97.");

            //get the table stream
            if (this.FIB.fWhichTblStm)
                this.TableStream = reader.GetStream("1Table");
            else
                this.TableStream = reader.GetStream("0Table");

            //get the data stream
            try
            {
                this.DataStream = reader.GetStream("Data");
            }
            catch (StreamNotFoundException)
            {
                this.DataStream = null;
            }

            //parse the stylesheet
            this.Styles = new StyleSheet(this.FIB, this.TableStream, this.DataStream);

            //read font table
            this.FontTable = new FontTable(this.TableStream, this.FIB);

            //parse all PAPX and build the dictionary
            this.AllPapx = new Dictionary<Int32, ParagraphPropertyExceptions>();
            List<FormattedDiskPagePAPX> allPapxFkps = FormattedDiskPagePAPX.GetAllPAPXFKPs(FIB, WordDocumentStream, TableStream, DataStream);
            for (int i=0; i < allPapxFkps.Count; i++)
            {
                for (int j = 0; j < allPapxFkps[i].grppapx.Length; j++)
                {
                    this.AllPapx.Add(allPapxFkps[i].rgfc[j], allPapxFkps[i].grppapx[j]);
                }
            }

            //parse the piece table and construct a list that contains all chars
            this.PieceTable = new PieceTable(this.FIB, this.TableStream);
            List<char> allChars = this.PieceTable.GetChars(this.FIB.fcMin, this.FIB.fcMac, this.WordDocumentStream);

            //split the chars into the subdocuments
            this.Text = allChars.GetRange(0, FIB.ccpText);
            this.Footnotes = allChars.GetRange(FIB.ccpText, FIB.ccpFtn);
            this.Headers = allChars.GetRange(FIB.ccpText + FIB.ccpFtn, FIB.ccpHdr);
            this.Annotations = allChars.GetRange(FIB.ccpText + FIB.ccpFtn + FIB.ccpHdr, FIB.ccpAtn);
            this.Endnotes = allChars.GetRange(FIB.ccpText + FIB.ccpFtn + FIB.ccpHdr + FIB.ccpAtn, FIB.ccpEdn);
            this.Textboxes = allChars.GetRange(FIB.ccpText + FIB.ccpFtn + FIB.ccpHdr + FIB.ccpAtn + FIB.ccpEdn, FIB.ccpTxbx);
            this.HeaderTextboxes = allChars.GetRange(FIB.ccpText + FIB.ccpFtn + FIB.ccpHdr + FIB.ccpAtn + FIB.ccpEdn + FIB.ccpTxbx, FIB.ccpHdrTxbx);
        }

        /// <summary>
        /// Finds the PAPX that is valid for the given FC.
        /// </summary>
        /// <param name="fc"></param>
        /// <returns></returns>
        public ParagraphPropertyExceptions FindValidPapx(Int32 fc)
        {
            ParagraphPropertyExceptions ret = null;

            while(ret == null)
            {
                try
                {
                    ret = AllPapx[fc];
                }
                catch (KeyNotFoundException){
                    fc--;
                }
            }

            return ret;
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WordDocument>)mapping).Apply(this);
        }

        #endregion
    }
}
