using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using System.IO;
using System.Xml;
using Microsoft.Office.Interop.Word;

namespace UnitTests
{
    [TestFixture]
    public class DocFileFormatTests
    {
        private Application word2007 = null;
        private List<FileInfo> files;
        private Object confirmConversions = Type.Missing;
        private Object readOnly = true;
        private Object addToRecentFiles = Type.Missing;
        private Object passwordDocument = Type.Missing;
        private Object passwordTemplate = Type.Missing;
        private Object revert = Type.Missing;
        private Object writePasswordDocument = Type.Missing;
        private Object writePasswordTemplate = Type.Missing;
        private Object format = Type.Missing;
        private Object encoding = Type.Missing;
        private Object visible = Type.Missing;
        private Object openConflictDocument = Type.Missing;
        private Object openAndRepair = Type.Missing;
        private Object documentDirection = Type.Missing;
        private Object noEncodingDialog = Type.Missing;
        private Object saveChanges = Type.Missing;
        private Object originalFormat = Type.Missing;
        private Object routeDocument = Type.Missing;

        [TestFixtureSetUp]
        public void SetUp()
        {
            //read the config
            FileStream fs = new FileStream("Config.xml", FileMode.Open);
            XmlDocument config = new XmlDocument();
            config.Load(fs);
            fs.Close();

            //read the inputfiles
            this.files = new List<FileInfo>();
            foreach (XmlNode fileNode in config.SelectNodes("input-files/file"))
            {
                this.files.Add(new FileInfo(fileNode.Attributes["path"].Value));
            }

            //start the application
            this.word2007 = new Application();
        }

        [TestFixtureTearDown]
        public void TearDown()
        {
            this.word2007.Quit(
                ref saveChanges,
                ref originalFormat,
                ref routeDocument);
        }

        private Document loadDocument(Object filename)
        {
            return this.word2007.Documents.Open(
                ref filename,
                ref confirmConversions,
                ref readOnly,
                ref addToRecentFiles,
                ref passwordDocument,
                ref passwordTemplate,
                ref revert,
                ref writePasswordDocument,
                ref writePasswordTemplate,
                ref format,
                ref encoding,
                ref visible,
                ref openConflictDocument,
                ref openAndRepair,
                ref documentDirection,
                ref noEncodingDialog);
        }

        /// <summary>
        /// Tests if the inputfile is parsable
        /// </summary>
        [Test]
        public void TestParseability()
        {
            foreach (FileInfo inputFile in this.files)
            {
                try
                {
                    StructuredStorageReader reader = new StructuredStorageReader(inputFile.FullName);
                    WordDocument doc = new WordDocument(reader);
                    Console.WriteLine("TestParseability: \"" + inputFile.FullName + "\"");
                }
                catch (Exception)
                {
                    Console.Error.WriteLine("TestParseability: \"" + inputFile.FullName + "\"");
                    Assert.Fail();
                }
            }
        }

        [Test]
        public void TestComments()
        {
            foreach (FileInfo inputFile in this.files)
            {
                Document omDoc = loadDocument(inputFile.FullName);
                WordDocument dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));

                int dffCommentCount = dffDoc.AnnotationsReferencePlex.Elements.Count;
                int omCommentCount = omDoc.Comments.Count;
                Comment omFirstComment = null;
                AnnotationReferenceDescriptor dffFirstComment = null;
                AnnotationReferenceDescriptorExtra dffFirstCommentExtra = null;

                if (dffCommentCount > 0 && omCommentCount > 0)
                {
                    omFirstComment = omDoc.Comments[1];
                    dffFirstComment = (AnnotationReferenceDescriptor)dffDoc.AnnotationsReferencePlex.Elements[1];
                    dffFirstCommentExtra = dffDoc.AnnotationReferenceExtraTable[1];
                }
                
                omDoc.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                try
                {
                    //test comment count
                    Assert.AreEqual(omCommentCount, dffCommentCount);

                    //compare the first comments
                    if (omFirstComment != null && dffFirstComment != null)
                    {
                        //compare initials
                        Assert.AreEqual(omFirstComment.Initial, dffFirstComment.UserInitials);

                        //compare dates
                        if (dffFirstCommentExtra != null)
                        {
                            Assert.AreEqual(omFirstComment.Date.Day, dffFirstCommentExtra.Date.dom);
                            Assert.AreEqual(omFirstComment.Date.Month, dffFirstCommentExtra.Date.mon);
                            Assert.AreEqual(omFirstComment.Date.Year, dffFirstCommentExtra.Date.yr);
                        }
                    }
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }

        
        /// <summary>
        /// 
        /// </summary>
        [Test]
        public void TestStyleCount()
        {
            foreach (FileInfo inputFile in this.files)
            {
                Document omDoc = loadDocument(inputFile.FullName);
                WordDocument dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));

                //count the styles of the MS document
                int omStyleCount = 0;
                foreach (Style s in omDoc.Styles)
                {
                    if (s.InUse)
                    {
                        omStyleCount++;
                    }
                }

                //count the styles of the DFF document
                int dffStyleCount = 0;
                foreach (StyleSheetDescription s in dffDoc.Styles.Styles)
                {
                    if (s != null && s.sti != StyleSheetDescription.StyleIdentifier.User)
                    {
                        dffStyleCount++;
                    }
                }


                omDoc.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                try
                {
                    Assert.AreEqual(omStyleCount, dffStyleCount);
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }

        [Test]
        public void TestMainText()
        {
            foreach (FileInfo inputFile in this.files)
            {
                Document word2007Doc = loadDocument(inputFile.FullName);
                WordDocument dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));

                //the text that is delivered by the object model and by DFF are always different 
                //because the objetc model replaces some control characters.
                //for this reason we remove all control characters

                StringBuilder dffMainText = new StringBuilder();
                foreach(char c in word2007Doc.StoryRanges[WdStoryType.wdMainTextStory].Text.ToCharArray())
                {
                    if((int)c >= 0x21)
                    {
                        dffMainText.Append(c);
                    }
                }

                StringBuilder omMainText = new StringBuilder();
                //dffDoc.Text.GetRange(0, dffDoc.FIB.ccpText)
                foreach (char c in dffDoc.Text.ToArray())
                {
                    if((int)c >= 0x21)
                    {
                        omMainText.Append(c);
                    }
                }

                word2007Doc.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                try
                {
                    Assert.AreEqual(omMainText.ToString(), dffMainText.ToString());
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }

        [Test]
        public void TestCharacterCount()
        {
            foreach (FileInfo inputFile in this.files)
            {
                Document omDoc = loadDocument(inputFile.FullName);
                WordDocument dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));

                int omCount = omDoc.Characters.Count;
                int dffCount = dffDoc.Text.GetRange(0, dffDoc.FIB.ccpText).Count;

                omDoc.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                try
                {
                    Assert.AreEqual(omCount, dffCount);
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }

    }
}
