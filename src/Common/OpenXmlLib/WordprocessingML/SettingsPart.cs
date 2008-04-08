using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class SettingsPart : UniqueOpenXmlPart
    {
        internal SettingsPart(OpenXmlPartContainer parent)
            : base(parent)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Comments; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Comments; }
        }

        public override string TargetName { get { return "settings"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
