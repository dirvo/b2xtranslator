using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML
{
    public class SlideMasterPart : ContentPart
    {
        public List<SlideLayoutPart> SlideLayoutParts = new List<SlideLayoutPart>();
        protected int _slideLayoutCounter;

        public SlideMasterPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        } 
        
        public override string ContentType
        {
            get { return PresentationMLContentTypes.SlideMaster; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.SlideMaster; }
        }

        public override string TargetName { get { return "slideMaster" + this.PartIndex; } }
        public override string TargetDirectory { get { return "slideMasters"; } }

        public SlideLayoutPart AddSlideLayoutPart()
        {
            SlideLayoutPart part = new SlideLayoutPart(this, ++_slideLayoutCounter);
            this.SlideLayoutParts.Add(part);
            return this.AddPart(part);
        }

        public ThemePart AddThemePart()
        {
            return this.AddPart(new ThemePart(this));
        }
    }
}
