using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.WordprocessingML;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.IO;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.WordprocessingMLMapping;

namespace DocTranslatorTest
{
    class Program
    {
        private static StorageReader reader;
        private static VirtualStream wordDocumentStream, tableStream;
        private static FileInformationBlock fib;
        private static string file, method;

        static void Main(string[] args)
        {
            //try
            //{
                //parse arguments
                parseArgs(args);

                reader = new StorageReader(file);

                //get the "WordDocument" stream
                wordDocumentStream = reader.GetStream("WordDocument");

                //parse the FIB
                fib = new FileInformationBlock(wordDocumentStream);

                //starting
                if (!fib.fComplex)
                {
                    //get the tablestream
                    if (fib.fWhichTblStm)
                        tableStream = reader.GetStream("1Table");
                    else
                        tableStream = reader.GetStream("0Table");

                    using (WordprocessingDocument doc = WordprocessingDocument.Create(file + "x", WordprocessingDocumentType.Document))
                    {
                        XmlWriterSettings xws = new XmlWriterSettings();
                        xws.OmitXmlDeclaration = false;
                        xws.CloseOutput = true;
                        xws.Encoding = Encoding.UTF8;
                        xws.Indent = true;
                        xws.ConformanceLevel = ConformanceLevel.Document;

                        Stream stream = doc.StyleDefinitionsPart.GetStream();
                        XmlWriter writer = XmlWriter.Create(doc.StyleDefinitionsPart.GetStream(), xws);
                        //writer = new XmlTextWriter(stream, Encoding.UTF8);

                        ////writer.WriteStartDocument();
                        ////writer.WriteStartElement("styles", OpenXmlNamespaces.WordprocessingML);

                        ////writer.WriteEndElement();
                        ////writer.WriteEndDocument();
                        
                        ////writer.Flush();
                        ////AddContent(writer);
                        StyleSheetMapping mapping = new StyleSheetMapping(writer);
                        StyleSheet stsh = new StyleSheet(fib, tableStream);

                        stsh.Convert(mapping);
                        writer.Flush();
                    }
                }
                else
                {
                    Console.WriteLine(file + " has been fast-saved. This format is currently not supported.");
                }

                reader.Close();
            //}
            //catch (ArgumentException ae)
            //{
            //    printUsage();
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.ToString());
            //}
            //finally
            //{
            //    if (reader != null)
            //        reader.Close();
            //}
        }

        public static void AddContent(XmlWriter writer)
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("styles", OpenXmlNamespaces.WordprocessingML);

            writer.WriteEndElement();
            writer.WriteEndDocument();

            writer.Flush();
        }

        /// <summary>
        /// Parses the arguments
        /// </summary>
        /// <param name="args"></param>
        private static void parseArgs(string[] args)
        {
            try
            {
                file = args[0];
                //FileInfo fi = new FileInfo(file);
                //method = args[1];
            }
            catch (Exception)
            {
                throw new ArgumentException();
            }
        }

        /// <summary>
        /// Prints the usage of the tool
        /// </summary>
        private static void printUsage()
        {
            Console.WriteLine("Usage: DocTranslatorTest.exe {filename}");
        }
    }
}
