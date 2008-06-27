using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using System.IO;
using System.Xml;
using System.Reflection;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public static class Utils
    {
        private static readonly double MC_PER_EMU = 1587.5;

        public static Int32 MasterCoordToEMU(Int32 mc)
        {
            return (Int32) (mc * MC_PER_EMU);
        }

        public static Int32 EMUToMasterCoord(Int32 emu)
        {
            return (Int32) (emu / MC_PER_EMU);
        }

        public static XmlDocument GetDefaultDocument(string filename)
        {
            Assembly a = Assembly.GetExecutingAssembly();
            Stream s = a.GetManifestResourceStream(String.Format("{0}.Defaults.{1}.xml",
                typeof(Utils).Namespace, filename));

            XmlDocument doc = new XmlDocument();
            doc.Load(s);
            return doc;
        }

        public static string SlideSizeTypeToXMLValue(SlideSizeType sst)
        {
            // OOXML Spec § 4.8.22
            switch (sst)
            {
                case SlideSizeType.A4Paper:
                    return "A4";

                case SlideSizeType.Banner:
                    return "banner";

                case SlideSizeType.Custom:
                    return "custom";

                case SlideSizeType.LetterSizedPaper:
                    return "letter";

                case SlideSizeType.OnScreen:
                    return "screen4x3";

                case SlideSizeType.Overhead:
                    return "overhead";

                case SlideSizeType.Size35mm:
                    return "35mm";

                default:
                    throw new NotImplementedException(
                        String.Format("Can't convert slide size type {0} to XML value", sst));
            }
        }

        public static string PlaceholderIdToXMLValue(PlaceholderId pid)
        {
            switch (pid)
            {
                case PlaceholderId.MasterTitle:
                    return "title";

                case PlaceholderId.MasterBody:
                    return "body";

                case PlaceholderId.MasterDate:
                    return "dt";

                case PlaceholderId.MasterSlideNumber:
                    return "sldNum";

                case PlaceholderId.MasterFooter:
                    return "ftr";

                case PlaceholderId.Title:
                    return "title";

                case PlaceholderId.Body:
                    return "body";

                case PlaceholderId.CenteredTitle:
                    return "ctrTitle";

                case PlaceholderId.Subtitle:
                    return "subTitle";

                default:
                    return null;
            }
        }
    }
}
