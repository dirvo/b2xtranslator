using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public class VbaDataPart: ContentPart
    {
        internal VbaDataPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return OpenXmlContentTypes.VbaData; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.VbaData; }
        }

        public override string TargetName { get { return "vbaData"; } }
        public override string TargetExt { get { return ".xml"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
