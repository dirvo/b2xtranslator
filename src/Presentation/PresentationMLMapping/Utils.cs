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

        public static string PlaceholderIdToXMLValue(PlaceholderEnum pid)
        {
            switch (pid)
            {
                case PlaceholderEnum.MasterDate:
                    return "dt";

                case PlaceholderEnum.MasterSlideNumber:
                    return "sldNum";

                case PlaceholderEnum.MasterFooter:
                    return "ftr";

                case PlaceholderEnum.MasterTitle:
                case PlaceholderEnum.Title:
                    return "title";

                case PlaceholderEnum.MasterBody:
                case PlaceholderEnum.Body:
                    return "body";

                case PlaceholderEnum.MasterCenteredTitle:
                case PlaceholderEnum.CenteredTitle:
                    return "ctrTitle";

                case PlaceholderEnum.MasterSubtitle:
                case PlaceholderEnum.Subtitle:
                    return "subTitle";

                case PlaceholderEnum.ClipArt:
                    return "clipArt";

                case PlaceholderEnum.Graph:
                    return "chart";

                case PlaceholderEnum.OrganizationChart:
                    return "dgm";

                case PlaceholderEnum.MediaClip:
                    return "media";

                case PlaceholderEnum.Table:
                    return "tbl";

                default:
                    throw new NotImplementedException("Don't know how to map placeholder id " + pid);
            }
        }

        public static string SlideLayoutTypeToFilename(SlideLayoutType type, PlaceholderEnum[] placeholderTypes)
        {
            switch (type)
            {
                case SlideLayoutType.BigObject:
                    return "objOnly";

                case SlideLayoutType.Blank:
                    return "blank";

                case SlideLayoutType.FourObjects:
                    return "fourObj";

                case SlideLayoutType.TitleAndBody:
                    {
                        PlaceholderEnum body = placeholderTypes[1];

                        if (body == PlaceholderEnum.Table)
                        {
                            return "tbl";
                        }
                        else if (body == PlaceholderEnum.OrganizationChart)
                        {
                            return "dgm";
                        }
                        else if (body == PlaceholderEnum.Graph)
                        {
                            return "chart";
                        }
                        else
                        {
                            return "obj";
                        }
                    }

                case SlideLayoutType.TitleOnly:
                    return "titleOnly";

                case SlideLayoutType.TitleSlide:
                    return "title";

                case SlideLayoutType.TwoColumnsAndTitle:
                    {
                        PlaceholderEnum leftType = placeholderTypes[1];
                        PlaceholderEnum rightType = placeholderTypes[2];

                        if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.Object)
                        {
                            return "txAndObj";
                        }
                        else if (leftType == PlaceholderEnum.Object && rightType == PlaceholderEnum.Body)
                        {
                            return "objAndTx";
                        }
                        else if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.ClipArt)
                        {
                            return "txAndClipArt";
                        }
                        else if (leftType == PlaceholderEnum.ClipArt && rightType == PlaceholderEnum.Body)
                        {
                            return "clipArtAndTx";
                        }
                        else if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.Graph)
                        {
                            return "txAndChart";
                        }
                        else if (leftType == PlaceholderEnum.Graph && rightType == PlaceholderEnum.Body)
                        {
                            return "chartAndTx";
                        }
                        else if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.MediaClip)
                        {
                            return "txAndMedia";
                        }
                        else if (leftType == PlaceholderEnum.MediaClip && rightType == PlaceholderEnum.Body)
                        {
                            return "mediaAndTx";
                        }
                        else
                        {
                            return "twoObj";
                        }
                    }

                case SlideLayoutType.TwoColumnsLeftTwoRows:
                    {
                        PlaceholderEnum rightType = placeholderTypes[2];

                        if (rightType == PlaceholderEnum.Object)
                        {
                            return "twoObjAndObj";
                        }
                        else if (rightType == PlaceholderEnum.Body)
                        {
                            return "twoObjAndTx";
                        }
                        else
                        {
                            throw new NotImplementedException(String.Format(
                                "Don't know how to map TwoColumnLeftTwoRows with rightType = {0}",
                                rightType
                            ));
                        }
                    }

                case SlideLayoutType.TwoColumnsRightTwoRows:
                    {
                        PlaceholderEnum leftType = placeholderTypes[1];

                        if (leftType == PlaceholderEnum.Object)
                        {
                            return "objAndTwoObj";
                        }
                        else if (leftType == PlaceholderEnum.Body)
                        {
                            return "txAndTwoObj";
                        }
                        else
                        {
                            throw new NotImplementedException(String.Format(
                                "Don't know how to map TwoColumnRightTwoRows with leftType = {0}",
                                leftType
                            ));
                        }
                    }

                case SlideLayoutType.TwoRowsAndTitle:
                    {
                        PlaceholderEnum topType = placeholderTypes[1];
                        PlaceholderEnum bottomType = placeholderTypes[2];

                        if (topType == PlaceholderEnum.Body && bottomType == PlaceholderEnum.Object)
                        {
                            return "txOverObj";
                        }
                        else if (topType == PlaceholderEnum.Object && bottomType == PlaceholderEnum.Body)
                        {
                            return "objOverTx";
                        }
                        else
                        {
                            throw new NotImplementedException(String.Format(
                                "Don't know how to map TwoRowsAndTitle with topType = {0} and bottomType = {1}",
                                topType, bottomType
                            ));
                        }
                    }

                case SlideLayoutType.TwoRowsTopTwoColumns:
                    return "twoObjOverTx";

                default:
                    throw new NotImplementedException("Don't know how to map slide layout type " + type); 
            }
        }
    }
}
