using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    class ExtractorException: Exception 
    {
        /// <summary>
        /// some public static attributes 
        /// </summary>
        public const string NULLPOINTEREXCEPTION = "Null pointer exception!!";
        public const string NOFILEFOUNDEXCEPTION = "No file found!!"; 


        /// <summary>
        /// Overridden ctor 
        /// </summary>
        public ExtractorException()
        {
        }

        /// <summary>
        /// Overridden ctor
        /// </summary>
        /// <param name="message">The exception message</param>
        public ExtractorException(string message)
        : base(message)
        {
        }

        /// <summary>
        /// Overridden ctor
        /// </summary>
        /// <param name="message">The exception message</param>
        /// <param name="inner"></param>
        public ExtractorException(string message, Exception inner)
        : base(message, inner)
        {
        }
    }
}
