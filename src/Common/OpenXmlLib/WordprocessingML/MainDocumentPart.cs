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

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class MainDocumentPart : UniqueOpenXmlPart
    {
        protected StyleDefinitionsPart _styleDefinitionsPart;
        protected FontTablePart _fontTablePart;
        protected NumberingDefinitionsPart _numberingDefinitionsPart;

        protected int _imagePartCount = 0;
        
        public MainDocumentPart(OpenXmlPartContainer parent)
            : base(parent)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.MainDocument; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.OfficeDocument; }
        }

        public override string TargetName { get { return "document"; } }
        public override string TargetDirectory { get { return "word"; } }


        public StyleDefinitionsPart AddStyleDefinitionsPart()
        {
            _styleDefinitionsPart = new StyleDefinitionsPart(this);
            return this.AddPart(_styleDefinitionsPart);
        }

        public StyleDefinitionsPart StyleDefinitionsPart
        {
            get { return _styleDefinitionsPart; }
        }

        public FontTablePart AddFontTablePart()
        {
            _fontTablePart = new FontTablePart(this);
            return this.AddPart(_fontTablePart);
        }

        public FontTablePart FontTablePart
        {
            get { return _fontTablePart; }
        }

        public NumberingDefinitionsPart AddNumberingDefinitionsPart()
        {
            _numberingDefinitionsPart = new NumberingDefinitionsPart(this);
            return this.AddPart(_numberingDefinitionsPart);
        }

        public NumberingDefinitionsPart NumberingDefinitionsPart
        {
            get { return _numberingDefinitionsPart; }
        }

        public ImagePart AddImagePart(ImagePartType type)
        {
            return this.AddPart(new ImagePart(type, this, ++_imagePartCount));
        }
    }
}
