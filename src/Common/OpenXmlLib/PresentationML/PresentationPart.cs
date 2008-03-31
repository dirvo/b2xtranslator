using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML
{
    public class PresentationPart : OpenXmlPart, IUniquePart
    {
        public PresentationPart(OpenXmlPartContainer parent)
            : base(parent)
        {
        }

        public override string ContentType
        {
            get { return PresentationMLContentTypes.Presentation; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.OfficeDocument; }
        }

        public override string TargetName { get { return "presentation"; } }
        public override string TargetDirectory { get { return "ppt"; } }
    }
}
