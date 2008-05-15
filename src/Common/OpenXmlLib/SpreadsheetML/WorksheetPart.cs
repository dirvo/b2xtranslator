using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.Spreadsheet
{

    public class WorksheetPart : OpenXmlPart
    {
        private int WorksheetNumber; 

        public WorksheetPart(OpenXmlPartContainer parent, int WorksheetNumber)
            : base(parent, WorksheetNumber)
        {
            this.WorksheetNumber = WorksheetNumber; 
        }


        public override string ContentType
        {
            get { return SpreadsheetMLContentTypes.Worksheet; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.WorkSheet; }
        }

        public override string TargetName { get { return "sheet" + this.WorksheetNumber.ToString(); } }
        public override string TargetDirectory { get { return "worksheets"; } }
    }
}
