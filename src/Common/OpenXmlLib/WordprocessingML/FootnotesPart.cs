using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class FootnotesPart : UniqueOpenXmlPart
    {
        public FootnotesPart(OpenXmlPartContainer parent)
            : base(parent)
        {
        }
        
        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Footnotes; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Footnotes; }
        }

        public override string TargetName { get { return "footnotes"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
