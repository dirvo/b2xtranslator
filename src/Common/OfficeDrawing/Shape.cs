using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(0xF00A)]
    public class Shape : Record
    {
        public Int32 spid;

        /// <summary>
        /// This shape is a group shape 
        /// </summary>
        public bool fGroup;

        /// <summary>
        /// Not a top-level shape 
        /// </summary>
        public bool fChild;

        /// <summary>
        /// This is the topmost group shape.<br/>
        /// Exactly one of these per drawing. 
        /// </summary>
        public bool fPatriarch; 

        /// <summary>
        /// The shape has been deleted 
        /// </summary>
        public bool fDeleted;

        /// <summary>
        /// The shape is an OLE object 
        /// </summary>
        public bool fOleShape;

        /// <summary>
        /// Shape has a hspMaster property 
        /// </summary>
        public bool fHaveMaster;

        /// <summary>
        /// Shape is flipped horizontally 
        /// </summary>
        public bool fFlipH;

        /// <summary>
        /// Shape is flipped vertically 
        /// </summary>
        public bool fFlipV;

        /// <summary>
        /// Connector type of shape 
        /// </summary>
        public bool fConnector;

        /// <summary>
        /// Shape has an anchor of some kind 
        /// </summary>
        public bool fHaveAnchor;

        /// <summary>
        /// Background shape 
        /// </summary>
        public bool fBackground;

        /// <summary>
        /// Shape has a shape type property
        /// </summary>
        public bool fHaveSpt;

        public Shape(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.spid = this.Reader.ReadInt32();

            UInt32 flag = this.Reader.ReadUInt32();
            this.fGroup = Utils.BitmaskToBool(flag, 0x1);
            this.fChild = Utils.BitmaskToBool(flag, 0x2);
            this.fPatriarch = Utils.BitmaskToBool(flag, 0x4);
            this.fDeleted = Utils.BitmaskToBool(flag, 0x8);
            this.fOleShape = Utils.BitmaskToBool(flag, 0x10);
            this.fHaveMaster = Utils.BitmaskToBool(flag, 0x20);
            this.fFlipH = Utils.BitmaskToBool(flag, 0x40);
            this.fFlipV = Utils.BitmaskToBool(flag, 0x80);
            this.fConnector = Utils.BitmaskToBool(flag, 0x100);
            this.fHaveAnchor = Utils.BitmaskToBool(flag, 0x200);
            this.fBackground = Utils.BitmaskToBool(flag, 0x400);
            this.fHaveSpt = Utils.BitmaskToBool(flag, 0x800);
        }

        public enum ShapeType
        {
            msosptMin = 0,
            msosptNotPrimitive = msosptMin,
            msosptRectangle = 1,
            msosptRoundRectangle = 2,
            msosptEllipse = 3,
            msosptDiamond = 4,
            msosptIsocelesTriangle = 5,
            msosptRightTriangle = 6,
            msosptParallelogram = 7,
            msosptTrapezoid = 8,
            msosptHexagon = 9,
            msosptOctagon = 10,
            msosptPlus = 11,
            msosptStar = 12,
            msosptArrow = 13,
            msosptThickArrow = 14,
            msosptHomePlate = 15,
            msosptCube = 16,
            msosptBalloon = 17,
            msosptSeal = 18,
            msosptArc = 19,
            msosptLine = 20,
            msosptPlaque = 21,
            msosptCan = 22,
            msosptDonut = 23,
            msosptTextSimple = 24,
            msosptTextOctagon = 25,
            msosptTextHexagon = 26,
            msosptTextCurve = 27,
            msosptTextWave = 28,
            msosptTextRing = 29,
            msosptTextOnCurve = 30,
            msosptTextOnRing = 31,
            msosptStraightConnector1 = 32,
            msosptBentConnector2 = 33,
            msosptBentConnector3 = 34,
            msosptBentConnector4 = 35,
            msosptBentConnector5 = 36,
            msosptCurvedConnector2 = 37,
            msosptCurvedConnector3 = 38,
            msosptCurvedConnector4 = 39,
            msosptCurvedConnector5 = 40,
            msosptCallout1 = 41,
            msosptCallout2 = 42,
            msosptCallout3 = 43,
            msosptAccentCallout1 = 44,
            msosptAccentCallout2 = 45,
            msosptAccentCallout3 = 46,
            msosptBorderCallout1 = 47,
            msosptBorderCallout2 = 48,
            msosptBorderCallout3 = 49,
            msosptAccentBorderCallout1 = 50,
            msosptAccentBorderCallout2 = 51,
            msosptAccentBorderCallout3 = 52,
            msosptRibbon = 53,
            msosptRibbon2 = 54,
            msosptChevron = 55,
            msosptPentagon = 56,
            msosptNoSmoking = 57,
            msosptSeal8 = 58,
            msosptSeal16 = 59,
            msosptSeal32 = 60,
            msosptWedgeRectCallout = 61,
            msosptWedgeRRectCallout = 62,
            msosptWedgeEllipseCallout = 63,
            msosptWave = 64,
            msosptFoldedCorner = 65,
            msosptLeftArrow = 66,
            msosptDownArrow = 67,
            msosptUpArrow = 68,
            msosptLeftRightArrow = 69,
            msosptUpDownArrow = 70,
            msosptIrregularSeal1 = 71,
            msosptIrregularSeal2 = 72,
            msosptLightningBolt = 73,
            msosptHeart = 74,
            msosptPictureFrame = 75,
            msosptQuadArrow = 76,
            msosptLeftArrowCallout = 77,
            msosptRightArrowCallout = 78,
            msosptUpArrowCallout = 79,
            msosptDownArrowCallout = 80,
            msosptLeftRightArrowCallout = 81,
            msosptUpDownArrowCallout = 82,
            msosptQuadArrowCallout = 83,
            msosptBevel = 84,
            msosptLeftBracket = 85,
            msosptRightBracket = 86,
            msosptLeftBrace = 87,
            msosptRightBrace = 88,
            msosptLeftUpArrow = 89,
            msosptBentUpArrow = 90,
            msosptBentArrow = 91,
            msosptSeal24 = 92,
            msosptStripedRightArrow = 93,
            msosptNotchedRightArrow = 94,
            msosptBlockArc = 95,
            msosptSmileyFace = 96,
            msosptVerticalScroll = 97,
            msosptHorizontalScroll = 98,
            msosptCircularArrow = 99,
            msosptNotchedCircularArrow = 100,
            msosptUturnArrow = 101,
            msosptCurvedRightArrow = 102,
            msosptCurvedLeftArrow = 103,
            msosptCurvedUpArrow = 104,
            msosptCurvedDownArrow = 105,
            msosptCloudCallout = 106,
            msosptEllipseRibbon = 107,
            msosptEllipseRibbon2 = 108,
            msosptFlowChartProcess = 109,
            msosptFlowChartDecision = 110,
            msosptFlowChartInputOutput = 111,
            msosptFlowChartPredefinedProcess = 112,
            msosptFlowChartInternalStorage = 113,
            msosptFlowChartDocument = 114,
            msosptFlowChartMultidocument = 115,
            msosptFlowChartTerminator = 116,
            msosptFlowChartPreparation = 117,
            msosptFlowChartManualInput = 118,
            msosptFlowChartManualOperation = 119,
            msosptFlowChartConnector = 120,
            msosptFlowChartPunchedCard = 121,
            msosptFlowChartPunchedTape = 122,
            msosptFlowChartSummingJunction = 123,
            msosptFlowChartOr = 124,
            msosptFlowChartCollate = 125,
            msosptFlowChartSort = 126,
            msosptFlowChartExtract = 127,
            msosptFlowChartMerge = 128,
            msosptFlowChartOfflineStorage = 129,
            msosptFlowChartOnlineStorage = 130,
            msosptFlowChartMagneticTape = 131,
            msosptFlowChartMagneticDisk = 132,
            msosptFlowChartMagneticDrum = 133,
            msosptFlowChartDisplay = 134,
            msosptFlowChartDelay = 135,
            msosptTextPlainText = 136,
            msosptTextStop = 137,
            msosptTextTriangle = 138,
            msosptTextTriangleInverted = 139,
            msosptTextChevron = 140,
            msosptTextChevronInverted = 141,
            msosptTextRingInside = 142,
            msosptTextRingOutside = 143,
            msosptTextArchUpCurve = 144,
            msosptTextArchDownCurve = 145,
            msosptTextCircleCurve = 146,
            msosptTextButtonCurve = 147,
            msosptTextArchUpPour = 148,
            msosptTextArchDownPour = 149,
            msosptTextCirclePour = 150,
            msosptTextButtonPour = 151,
            msosptTextCurveUp = 152,
            msosptTextCurveDown = 153,
            msosptTextCascadeUp = 154,
            msosptTextCascadeDown = 155,
            msosptTextWave1 = 156,
            msosptTextWave2 = 157,
            msosptTextWave3 = 158,
            msosptTextWave4 = 159,
            msosptTextInflate = 160,
            msosptTextDeflate = 161,
            msosptTextInflateBottom = 162,
            msosptTextDeflateBottom = 163,
            msosptTextInflateTop = 164,
            msosptTextDeflateTop = 165,
            msosptTextDeflateInflate = 166,
            msosptTextDeflateInflateDeflate = 167,
            msosptTextFadeRight = 168,
            msosptTextFadeLeft = 169,
            msosptTextFadeUp = 170,
            msosptTextFadeDown = 171,
            msosptTextSlantUp = 172,
            msosptTextSlantDown = 173,
            msosptTextCanUp = 174,
            msosptTextCanDown = 175,
            msosptFlowChartAlternateProcess = 176,
            msosptFlowChartOffpageConnector = 177,
            msosptCallout90 = 178,
            msosptAccentCallout90 = 179,
            msosptBorderCallout90 = 180,
            msosptAccentBorderCallout90 = 181,
            msosptLeftRightUpArrow = 182,
            msosptSun = 183,
            msosptMoon = 184,
            msosptBracketPair = 185,
            msosptBracePair = 186,
            msosptSeal4 = 187,
            msosptDoubleWave = 188,
            msosptActionButtonBlank = 189,
            msosptActionButtonHome = 190,
            msosptActionButtonHelp = 191,
            msosptActionButtonInformation = 192,
            msosptActionButtonForwardNext = 193,
            msosptActionButtonBackPrevious = 194,
            msosptActionButtonEnd = 195,
            msosptActionButtonBeginning = 196,
            msosptActionButtonReturn = 197,
            msosptActionButtonDocument = 198,
            msosptActionButtonSound = 199,
            msosptActionButtonMovie = 200,
            msosptHostControl = 201,
            msosptTextBox = 202,
            msosptMax,
            msosptNil = 0x0FFF
        }
    }
}
