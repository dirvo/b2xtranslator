using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class KeyMapCustomizationsPart : ContentPart
    {
        private ToolbarsPart _toolbars;

        public KeyMapCustomizationsPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return MicrosoftWordContentTypes.KeyMapCustomization; }
        }

        public override string RelationshipType
        {
            get { return MicrosoftWordRelationshipTypes.KeyMapCustomizations; }
        }

        public override string TargetName { get { return "customizations"; } }
        public override string TargetDirectory { get { return ""; } }

        public ToolbarsPart ToolbarsPart
        {   
            get {
                if (_toolbars == null)
                {
                    _toolbars = new ToolbarsPart(this);
                    this.AddPart(_toolbars);
                }
                return _toolbars; 
            }
        }
	
    }
}
