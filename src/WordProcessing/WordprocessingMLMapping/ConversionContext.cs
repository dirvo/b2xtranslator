using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class ConversionContext
    {
        private WordprocessingDocument _docx;
        private Dictionary<Int32, SectionPropertyExceptions> _allSepx;
        private Dictionary<Int32, ParagraphPropertyExceptions> _allPapx;
        private XmlWriterSettings _writerSettings;
        private WordDocument _doc;

        /// <summary>
        /// The source of the conversion.
        /// </summary>
        public WordDocument Doc
        {
            get { return _doc; }
            set { _doc = value; }
        }

        /// <summary>
        /// This is the target of the conversion.<br/>
        /// The result will be written to the parts of this document.
        /// </summary>
        public WordprocessingDocument Docx
        {
            get { return _docx; }
            set { _docx = value; }
        }

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return _writerSettings; }
            set { _writerSettings = value; }
        }

        /// <summary>
        /// A dictionary that contains all SEPX of the document.<br/>
        /// The key is the CP at which sections ends.<br/>
        /// The value is the SEPX that formats the section.
        /// </summary>
        public Dictionary<Int32, SectionPropertyExceptions> AllSepx
        {
            get { return _allSepx; }
            set { _allSepx = value; }
        }   

        /// <summary>
        /// A dictionary that contains all PAPX of the document.<br/>
        /// The key is the FC at which the paragraph starts.<br/>
        /// The value is the PAPX that formats the paragraph.
        /// </summary>
        public Dictionary<Int32, ParagraphPropertyExceptions> AllPapx
        {
            get { return _allPapx; }
            set { _allPapx = value; }
        }

        /// <summary>
        /// A list thta contains all revision ids.
        /// </summary>
        public List<string> AllRsids;

        public ConversionContext(WordDocument doc)
        {
            this.Doc = doc;
            this.AllRsids = new List<string>();

            //build a dictionaries of all PAPX
            _allPapx = new Dictionary<Int32, ParagraphPropertyExceptions>();
            for (int i = 0; i < doc.AllPapxFkps.Count; i++)
            {
                for (int j = 0; j < doc.AllPapxFkps[i].grppapx.Length; j++)
                {
                    _allPapx.Add(doc.AllPapxFkps[i].rgfc[j], doc.AllPapxFkps[i].grppapx[j]);
                }
            }

            //build a dictionary of all SEPX
            _allSepx = new Dictionary<Int32, SectionPropertyExceptions>();
            for (int i = 0; i < doc.SectionTable.grpsepx.Length; i++)
            {
                _allSepx.Add(doc.SectionTable.rgfc[i + 1], doc.SectionTable.grpsepx[i]);
            }
        }

        /// <summary>
        /// Adds a new RSID to the list
        /// </summary>
        /// <param name="rsid"></param>
        public void AddRsid(string rsid)
        {
            if (!this.AllRsids.Contains(rsid))
                this.AllRsids.Add(rsid);
        }
    }
}
