using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML
{
    public class PresentationPart : ContentPart
    {
        public List<SlideMasterPart> SlideMasterParts = new List<SlideMasterPart>();
        protected int _slideMasterCounter = 0;
        protected int _slideCounter = 0;
        
        public PresentationPart(OpenXmlPartContainer parent)
            : base(parent, 0)
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

        public SlideMasterPart AddSlideMasterPart()
        {
            SlideMasterPart part = new SlideMasterPart(this, ++_slideMasterCounter);
            this.SlideMasterParts.Add(part);
            return this.AddPart(part);
        }

        public SlidePart AddSlidePart()
        {
            return this.AddPart(new SlidePart(this, ++_slideCounter));
        }
    }
}
