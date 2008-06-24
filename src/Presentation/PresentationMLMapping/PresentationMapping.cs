using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.PptFileFormat;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public abstract class PresentationMapping<T> :
        AbstractOpenXmlMapping,
        IMapping<T>
        where T : IVisitable
    {
        protected ConversionContext _ctx;
        public ContentPart targetPart;
        
        public PresentationMapping(ConversionContext ctx, ContentPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), ctx.WriterSettings))
        {
            _ctx = ctx;
            this.targetPart = targetPart;
        }

        public abstract void Apply(T mapElement);
    }
}
