using System;
using System.Collections.Generic;
using System.Text;

namespace UnitTests
{
    public class Program
    {
        public static void Main(string[] args)
        {
            DocFileFormatTests test = new DocFileFormatTests();
            test.SetUp();

            test.TestStyleCount();

            test.TearDown();
        }
    }
}
