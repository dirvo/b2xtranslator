using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class NumberingDefinitionsPart : UniqueOpenXmlPart
    {
        public NumberingDefinitionsPart(OpenXmlPartContainer parent)
            : base(parent)
        {
        }
        
        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Numbering; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Numbering; }
        }

        public override string TargetName { get { return "numbering"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
