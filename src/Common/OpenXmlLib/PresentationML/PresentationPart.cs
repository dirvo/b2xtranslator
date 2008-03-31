using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML
{
    public class PresentationPart : UniqueOpenXmlPart
    {
        protected int _slideLayoutCounter = 0;
        protected int _slideMasterCounter = 0;
        protected int _slideCounter = 0;
        
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

        public SlideLayoutPart AddSlideLayoutPart()
        {
            return this.AddPart(new SlideLayoutPart(this, ++_slideLayoutCounter));
        }

        public SlideMasterPart AddSlideMasterPart()
        {
            return this.AddPart(new SlideMasterPart(this, ++_slideMasterCounter));
        }

        public SlidePart AddSlidePart()
        {
            return this.AddPart(new SlidePart(this, ++_slideCounter));
        }
    }
}
