using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AuthorTable : List<string>
    {
        /// <summary>
        /// Parses the bytes to retrieve a AuthorTable
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public AuthorTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            //read the bytes
            //byte[] bytes = new byte[fib.lcbSttbfRMark];
            //tableStream.Read(bytes, 0, (int)fib.lcbSttbfRMark, (int)fib.fcSttbfRMark);

            int pos = 8;
            byte[] uniChar = new byte[2];
            StringBuilder name = new StringBuilder();
            while (pos < fib.lcbSttbfRMark)
            {
                tableStream.Read(uniChar, 0, 2, (int)(fib.fcSttbfRMark + pos));
                char cPos = Encoding.Unicode.GetString(uniChar).ToCharArray()[0];
                if ((int)cPos > 0x1F)
                {
                    name.Append(cPos);
                }
                else
                {
                    //there is a seperator that terminates this name
                    this.Add(name.ToString());
                    name = new StringBuilder();
                }
                pos+=2;
            }
            //add last name
            this.Add(name.ToString());
        }
    }
}
