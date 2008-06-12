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
        /// Contains all section descriptors
        /// </summary>
        public SectionTable SectionTable;

        /// <summary>
        /// Contains the names of all author who revised something in the document
        /// </summary>
        public AuthorTable AuthorTable;

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
        /// All text of the Word document
        /// </summary>
        public List<char> Text;

        /// <summary>
        /// The style sheet of the document
        /// </summary>
        public StyleSheet Styles;

        /// <summary>
        /// A list of all font names, used in the doucument
        /// </summary>
        public FontTable FontTable;

        /// <summary>
        /// A list that contains all formatting information of 
        /// the lists and numberings in the document.
        /// </summary>
        public ListTable ListTable;

        /// <summary>
        /// The drawing object table ....
        /// </summary>
        public DrawingObjectTable DrawingObjectTable;

        /// <summary>
        /// 
        /// </summary>
        public OfficeDrawingTable OfficeDrawingTable;

        public TextboxLinkTable TextboxLinkTable;

        /// <summary>
        /// 
        /// </summary>
        public OfficeDrawingTable OfficeDrawingTableHeader;

        /// <summary>
        /// The DocumentProperties of the word document
        /// </summary>
        public DocumentProperties DocumentProperties;

        /// <summary>
        /// A list that contains all overriding formatting information
        /// of the lists and numberings in the document.
        /// </summary>
        public ListFormatOverrideTable ListFormatOverrideTable;

        /// <summary>
        /// A list of all FKPs that contain PAPX
        /// </summary>
        public List<FormattedDiskPagePAPX> AllPapxFkps;

        /// <summary>
        /// A list of all FKPs that contain CHPX
        /// </summary>
        public List<FormattedDiskPageCHPX> AllChpxFkps;

        /// <summary>
        /// A table that contains the positions of the headers and footer in the text.
        /// </summary>
        public HeaderAndFooterTable HeaderAndFooterTable;

        public WordDocument(StructuredStorageFile reader)
        {
            this.WordDocumentStream = reader.GetStream("WordDocument");

            //parse FIB
            this.FIB = new FileInformationBlock(this.WordDocumentStream);

            if (this.FIB.nFib < 105)
                throw new UnspportedFileVersionException("DocFileFormat doesn't support Word versions older than Word 97.");

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

            VirtualStreamReader tablereader = new VirtualStreamReader(TableStream);
            byte[] fndRef = tablereader.ReadBytes(this.FIB.fcPlcffndRef, (int)this.FIB.lcbPlcffndRef);

            //parse properties
            this.DocumentProperties = new DocumentProperties(this.FIB, this.TableStream);

            //parse the stylesheet
            this.Styles = new StyleSheet(this.FIB, this.TableStream, this.DataStream);

            //read font table
            this.FontTable = new FontTable(this.FIB,this.TableStream);

            //read list table
            this.ListTable = new ListTable(this.FIB, this.TableStream);

            //read lfo table
            this.ListFormatOverrideTable = new ListFormatOverrideTable(this.FIB, this.TableStream);

            //parse the AuthorTable
            this.AuthorTable = new AuthorTable(this.FIB, this.TableStream);

            //read all PAPX FKPS
            this.AllPapxFkps = FormattedDiskPagePAPX.GetAllPAPXFKPs(this.FIB, this.WordDocumentStream, this.TableStream, this.DataStream);

            //read all CHPX FKPS
            this.AllChpxFkps = FormattedDiskPageCHPX.GetAllCHPXFKPs(this.FIB, this.WordDocumentStream, this.TableStream);

            //read section table
            this.SectionTable = new SectionTable(this.FIB, this.TableStream, this.WordDocumentStream);

            //read the DO table
            this.DrawingObjectTable = new DrawingObjectTable(this.FIB, this.TableStream);

            //read the OfficeDrawing table
            this.OfficeDrawingTable = new OfficeDrawingTable(this, OfficeDrawingTable.OfficeDrawingTableType.MainDocument);

            //read the OfficeDrawing table
            this.OfficeDrawingTableHeader = new OfficeDrawingTable(this, OfficeDrawingTable.OfficeDrawingTableType.Header);

            //read headers and footer table
            this.HeaderAndFooterTable = new HeaderAndFooterTable(this);

            this.TextboxLinkTable = new TextboxLinkTable(this.FIB, this.TableStream);

            //parse the piece table and construct a list that contains all chars
            this.PieceTable = new PieceTable(this.FIB, this.TableStream);
            this.Text = this.PieceTable.GetChars(this.FIB.fcMin, this.FIB.fcMac, this.WordDocumentStream);
        }

        /// <summary>
        /// Returns a list of all CHPX which are valid for the given FCs.
        /// </summary>
        /// <param name="fcMin">The lower boundary</param>
        /// <param name="fcMax">The upper boundary</param>
        /// <returns>The FCs</returns>
        public List<Int32> GetFileCharacterPositions(Int32 fcMin, Int32 fcMax)
        {
            List<Int32> list = new List<Int32>();

            for (int i = 0; i < this.AllChpxFkps.Count; i++ )
            {
                FormattedDiskPageCHPX fkp = this.AllChpxFkps[i];

                //if the last fc of this fkp is smaller the fcMin
                //this fkp is before the requested range
                if (fkp.rgfc[fkp.rgfc.Length - 1] < fcMin)
                {
                    continue;
                }

                //if the first fc of this fkp is larger the Max
                //this fkp is beyond the requested range
                if (fkp.rgfc[0] > fcMax)
                {
                    break;
                }

                //don't add the duplicated values of the FKP boundaries (Length-1)
                int max = fkp.rgfc.Length - 1;

                //last fkp? 
                //use full table
                if (i == (this.AllChpxFkps.Count-1))
                {
                    max = fkp.rgfc.Length;
                }

                for (int j = 0; j < max; j++)
                {
                    if (fkp.rgfc[j] < fcMin && fkp.rgfc[j + 1] > fcMin)
                    {
                        //this chpx starts before fcMin
                        list.Add(fkp.rgfc[j]);
                    }
                    else if (fkp.rgfc[j] >= fcMin && fkp.rgfc[j] < fcMax)
                    {
                        //this chpx is in the range
                        list.Add(fkp.rgfc[j]);
                    }
                }
            }

            return list;
        }


        /// <summary>
        /// Returnes a list of all CharacterPropertyExceptions which correspond to text 
        /// between the given boundaries.
        /// </summary>
        /// <param name="fcMin">The lower boundary</param>
        /// <param name="fcMax">The upper boundary</param>
        /// <returns>The FCs</returns>
        public List<CharacterPropertyExceptions> GetCharacterPropertyExceptions(Int32 fcMin, Int32 fcMax)
        {
            List<CharacterPropertyExceptions> list = new List<CharacterPropertyExceptions>();

            foreach(FormattedDiskPageCHPX fkp in this.AllChpxFkps)
            {
                //geht the CHPX
                for (int j = 0; j < fkp.grpchpx.Length; j++)
                {
                    if (fkp.rgfc[j] < fcMin && fkp.rgfc[j + 1] > fcMin)
                    {
                        //this chpx starts before fcMin
                        list.Add(fkp.grpchpx[j]);
                    }
                    else if (fkp.rgfc[j] >= fcMin && fkp.rgfc[j] < fcMax)
                    {
                        //this chpx is in the range
                        list.Add(fkp.grpchpx[j]);
                    }
                }
            }

            return list;
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WordDocument>)mapping).Apply(this);
        }

        #endregion
    }
}
