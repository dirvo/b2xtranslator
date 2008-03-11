using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class PropertyExceptions : IVisitable
    {
        /// <summary>
        /// A list of the sprms that encode the differences between 
        /// CHP for a character and the PAP for the paragraph style used.
        /// </summary>
        public List<SinglePropertyModifier> grpprl;

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<PropertyExceptions>)mapping).Apply(this);
        }

        #endregion

    }

}
