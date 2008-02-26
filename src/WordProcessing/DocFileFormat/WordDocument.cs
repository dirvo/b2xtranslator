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
        private StorageReader _reader;
        private VirtualStream _wordDocumentStream, _tableStream;
        private PieceTable _pieceTable;

        /// <summary>
        /// The file information block of the word document
        /// </summary>
        public FileInformationBlock FIB;

        /// <summary>
        /// The text part of the Word document
        /// </summary>
        public List<char> TextPart;

        /// <summary>
        /// The style sheet of the document
        /// </summary>
        public StyleSheet Styles;

        public WordDocument(string filename)
        {
            _reader = new StorageReader(filename);
            _wordDocumentStream = _reader.GetStream("WordDocument");

            //parse FIB
            this.FIB = new FileInformationBlock(_wordDocumentStream);

            //get the table stream
            if (this.FIB.fWhichTblStm)
                _tableStream = _reader.GetStream("1Table");
            else
                _tableStream = _reader.GetStream("0Table");

            //parse the stylesheet
            this.Styles = new StyleSheet(this.FIB, _tableStream);

            //parse the piece table and construct the text part
            TextPart = new List<char>();
            _pieceTable = new PieceTable(this.FIB, _tableStream);
            foreach (PieceDescriptor pcd in _pieceTable.Pieces)
            {
                int charCount = pcd.cpEnd - pcd.cpStart;
                int byteCount = charCount;
                if(pcd.encoding == Encoding.Unicode)
                {
                    byteCount *= 2;
                }

                //read the bytes of that piece
                byte[] bytesOfPiece = new byte[byteCount];
                _wordDocumentStream.Read(bytesOfPiece, bytesOfPiece.Length, pcd.fc);

                //encode it
                char[] chars = pcd.encoding.GetString(bytesOfPiece).ToCharArray();

                //append it
                foreach (char c in chars)
                {
                    TextPart.Add(c);
                }
            }

            _reader.Close();
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WordDocument>)mapping).Apply(this);
        }

        #endregion
    }
}
