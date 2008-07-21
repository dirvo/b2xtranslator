using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.PptFileFormat;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class ConversionContext
    {
        private PresentationDocument _pptx;
        private XmlWriterSettings _writerSettings;
        private PowerpointDocument _ppt;

        private Dictionary<UInt32, MasterMapping> MasterIdToMapping = new Dictionary<UInt32, MasterMapping>();

        /// <summary>
        /// The source of the conversion.
        /// </summary>
        public PowerpointDocument Ppt
        {
            get { return _ppt; }
            set { _ppt = value; }
        }

        /// <summary>
        /// This is the target of the conversion.<br/>
        /// The result will be written to the parts of this document.
        /// </summary>
        public PresentationDocument Pptx
        {
            get { return _pptx; }
            set { _pptx = value; }
        }

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return _writerSettings; }
            set { _writerSettings = value; }
        }

        public ConversionContext(PowerpointDocument ppt)
        {
            this.Ppt = ppt;
        }

        /// <summary>
        /// Registers a MasterMapping so it can later be looked up by its master ID.
        /// </summary>
        /// <param name="masterId">Master id with which to associate the MasterMapping.</param>
        /// <param name="mapping">MasterMapping to be registered.</param>
        public void RegisterMasterMapping(UInt32 masterId, MasterMapping mapping)
        {
            this.MasterIdToMapping[masterId] = mapping;
        }

        /// <summary>
        /// Returns the MasterMapping associated with the specified master ID.
        /// </summary>
        /// <param name="masterId">Master ID for which to find a MasterMapping.</param>
        /// <returns>Found MasterMapping or null if none was found.</returns>
        public MasterMapping GetMasterMappingByMasterId(UInt32 masterId)
        {
            return this.MasterIdToMapping[masterId];
        }

        /// <summary>
        /// Returns the MasterMapping associated with the specified master ID if it exists.
        /// Else a new MasterMapping is created.
        /// </summary>
        /// <param name="masterId">Master ID for which to find or create a MasterMapping.</param>
        /// <returns>Found or created MasterMapping.</returns>
        public MasterMapping GetOrCreateMasterMappingByMasterId(UInt32 masterId)
        {
            if (!this.MasterIdToMapping.ContainsKey(masterId))
                this.MasterIdToMapping[masterId] = new MasterMapping(this);

            return this.MasterIdToMapping[masterId];
        }

        protected Dictionary<UInt32, MasterLayoutManager> MasterIdToLayoutManager =
            new Dictionary<UInt32, MasterLayoutManager>();

        public MasterLayoutManager GetOrCreateLayoutManagerByMasterId(UInt32 masterId)
        {
            if (!this.MasterIdToLayoutManager.ContainsKey(masterId))
                this.MasterIdToLayoutManager[masterId] = new MasterLayoutManager(this, masterId);

            return this.MasterIdToLayoutManager[masterId];
        }
    }

    public class MasterLayoutManager
    {
        protected ConversionContext _ctx;
        protected UInt32 MasterId;

        /// <summary>
        /// PPT2007 layouts are stored inline with the master and
        /// have an instance id for associating them with slides.
        /// </summary>
        public Dictionary<UInt32, SlideLayoutPart> InstanceIdToLayoutPart =
            new Dictionary<UInt32, SlideLayoutPart>();

        /// <summary>
        /// Pre-PPT2007 layouts are specified in SSlideLayoutAtom
        /// as a SlideLayoutType integer value. Each SlideLayoutType
        /// can be mapped to a layout XML file.
        /// </summary>
        public Dictionary<UInt32, SlideLayoutPart> LayoutTypeToLayoutPart =
            new Dictionary<UInt32, SlideLayoutPart>();

        /// <summary>
        /// Pre-PPT2007 TitleMaster slides need to be converted to
        /// SlideLayoutParts for OOXML. The SlideLayoutParts for
        /// TitleMaster slides are stored in this dictionary.
        /// </summary>
        public Dictionary<UInt32, SlideLayoutPart> TitleMasterIdToLayoutPart =
            new Dictionary<UInt32, SlideLayoutPart>();

        public MasterLayoutManager(ConversionContext ctx, UInt32 masterId)
        {
            this._ctx = ctx;
            this.MasterId = masterId;
        }

        public List<SlideLayoutPart> GetAllLayoutParts()
        {
            List<SlideLayoutPart> result = new List<SlideLayoutPart>();

            result.AddRange(this.InstanceIdToLayoutPart.Values);
            result.AddRange(this.LayoutTypeToLayoutPart.Values);
            result.AddRange(this.TitleMasterIdToLayoutPart.Values);

            return result;
        }

        public SlideLayoutPart AddLayoutPartWithInstanceId(UInt32 instanceId)
        {
            SlideMasterPart masterPart = _ctx.GetOrCreateMasterMappingByMasterId(this.MasterId).MasterPart;
            SlideLayoutPart layoutPart = masterPart.AddSlideLayoutPart();

            this.InstanceIdToLayoutPart.Add(instanceId, layoutPart);
            return layoutPart;
        }

        public SlideLayoutPart GetLayoutPartByInstanceId(UInt32 instanceId)
        {
            return this.InstanceIdToLayoutPart[instanceId];
        }

        public SlideLayoutPart GetOrCreateLayoutPartByLayoutType(UInt32 type)
        {
            SlideMasterPart masterPart = _ctx.GetOrCreateMasterMappingByMasterId(this.MasterId).MasterPart;

            if (!this.LayoutTypeToLayoutPart.ContainsKey(type))
            {
                string layoutFile = Utils.SlideLayoutTypeToFilename((SlideLayoutType)type);
                XmlDocument slideLayoutDoc = Utils.GetDefaultDocument("slideLayouts." + layoutFile);

                SlideLayoutPart layoutPart = masterPart.AddSlideLayoutPart();
                slideLayoutDoc.WriteTo(layoutPart.XmlWriter);
                layoutPart.XmlWriter.Flush();

                this.LayoutTypeToLayoutPart.Add(type, layoutPart);
            }

            return this.LayoutTypeToLayoutPart[type];
        }

        public SlideLayoutPart GetOrCreateLayoutPartForTitleMasterId(UInt32 titleMasterId)
        {
            SlideMasterPart masterPart = _ctx.GetOrCreateMasterMappingByMasterId(this.MasterId).MasterPart;

            if (!this.TitleMasterIdToLayoutPart.ContainsKey(titleMasterId))
            {
                Slide titleMaster = _ctx.Ppt.FindMasterRecordById(titleMasterId);
                SlideLayoutPart layoutPart = masterPart.AddSlideLayoutPart();
                new TitleMasterMapping(_ctx, layoutPart).Apply(titleMaster);
                this.TitleMasterIdToLayoutPart.Add(titleMasterId, layoutPart);
            }

            return this.TitleMasterIdToLayoutPart[titleMasterId];
        }
    }
}
