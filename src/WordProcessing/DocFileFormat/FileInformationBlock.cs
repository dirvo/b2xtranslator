/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FileInformationBlock
    {
        /// <summary>
        /// Magic number
        /// </summary>
        public UInt16 wIdent;

        /// <summary>
        /// Product version written by
        /// </summary>
        public UInt16 nProduct;

        /// <summary>
        /// Language stamp
        /// </summary>
        public UInt16 Lid;

        /// <summary>
        /// 
        /// </summary>
        public Int16 pnNext;

        /// <summary>
        /// Set if this document is a template
        /// </summary>
        public bool fDot;

        /// <summary>
        /// Set if this document is a glossary
        /// </summary>
        public bool fGlsy;

        /// <summary>
        /// Set if this document is a comlex fast-saved format
        /// </summary>
        public bool fComplex;

        /// <summary>
        /// Set if this document is a template
        /// </summary>
        public bool fHasPic;

        /// <summary>
        /// Count of times file was quick saved
        /// </summary>
        public UInt16 cQuickSaves;

        /// <summary>
        /// Set if file is encrypted
        /// </summary>
        public bool fEncrypted;

        /// <summary>
        /// 
        /// </summary>
        public bool fWhichTblStm;

        /// <summary>
        /// Set when user has recommended that file be read read-only
        /// </summary>
        public bool fReadOnlyRecommended;

        /// <summary>
        /// Set when file owner has made the file write reserved
        /// </summary>
        public bool fWriteReservation;

        /// <summary>
        /// Set when using extended character set in file
        /// </summary>
        public bool fExtChar;

        /// <summary>
        /// REVIEW
        /// </summary>
        public bool fLoadOverwrite;

        /// <summary>
        /// REVIEW
        /// </summary>
        public bool fFarEast;

        /// <summary>
        /// REVIEW
        /// </summary>
        public bool fCrypto;

        /// <summary>
        /// This file format is compatible with readers that 
        /// understand nFib at or above this value.
        /// </summary>
        public UInt16 nFibBack;

        /// <summary>
        /// File encrypted Key, only valid if fEncrypted
        /// </summary>
        public Int32 lKey;

        /// <summary>
        /// Environment in which file was created
        /// </summary>
        public byte Envr;

        /// <summary>
        /// When true, this file was last saved in the Macintosh environment
        /// </summary>
        public bool fMac;

        /// <summary>
        /// 
        /// </summary>
        public bool fEmptySpecial;

        /// <summary>
        /// 
        /// </summary>
        public bool fLoadOverridePage;

        /// <summary>
        /// 
        /// </summary>
        public bool fFutureSavedUndo;

        /// <summary>
        /// 
        /// </summary>
        public bool fWord97Saved;

        /// <summary>
        /// Default extended character set id for text in document stream (overridden by chp.chse).<br/>
        /// 0 by default: characters in doc stream should be interpreted using the ANSI character set used by Windows.<br/>
        /// 256: characters in doc stream should be interpreted using the Macintosh character set.
        /// </summary>
        public UInt16 Chs;

        /// <summary>
        /// Default extended character set id for text in internal data structures.<br/>
        /// 0 by default: characters in internal data structures should be interpreted using the ANSI character set used by Windows.<br/>
        /// 256: characters in internal data structures should be interpreted using the Macintosh character set.
        /// </summary>
        public UInt16 chsTables;

        /// <summary>
        /// File offset of first character of text.<br/>
        /// In non-complex files a CP can be transformed into an FC by the following transformation:<br/>
        /// fc = cp + fib.fcMin
        /// </summary>
        public Int32 fcMin;

        /// <summary>
        /// File offset of last character of text in document text stream + 1
        /// </summary>
        public Int32 fcMac;

        /// <summary>
        /// Count of fields in the array of "shorts"
        /// </summary>
        public UInt16 Csw;

        /// <summary>
        /// Unique number identifying the file's creator.<br/>
        /// 0x6A62 is the creator ID for Word and is reserved.<br/>
        /// Other creators should choose a different value.
        /// </summary>
        public Int16 wMagicCreated;

        /// <summary>
        /// Identifies the file's last modifier
        /// </summary>
        public Int16 wMagicRevised;

        /// <summary>
        /// Private data
        /// </summary>
        public Int16 wMagicCreatedPrivate;

        /// <summary>
        /// Private data
        /// </summary>
        public Int16 wMagicRevisedPrivate;

        /// <summary>
        /// Language id if document was written by East Asian version of Word (i.e. FIB.fFarEast is on)
        /// </summary>
        public Int16 lidFE;

        /// <summary>
        /// Number of fields in the array of longs
        /// </summary>
        public UInt16 Clw;

        /// <summary>
        /// File offset of last byte written to file + 1
        /// </summary>
        public Int32 cbMac;

        /// <summary>
        /// Contains the build date of the creator.<br/>
        /// 10695 indicates the creator program was compiled on Jan 6, 1995.
        /// </summary>
        public Int32 lProductCreated;

        /// <summary>
        /// Contains the build date of the file's last modifier
        /// </summary>
        public Int32 lProductRevised;

        /// <summary>
        /// Length of main document text stream 1
        /// </summary>
        public Int32 ccpText;

        /// <summary>
        /// Length of footnote subdocument text stream
        /// </summary>
        public Int32 ccpFtn;

        /// <summary>
        /// Length of header subdocument text stream
        /// </summary>
        public Int32 ccpHdr;

        /// <summary>
        /// Length of macro subdocument text stream, which should now always be 0
        /// </summary>
        public Int32 ccpMcr;

        /// <summary>
        /// Length of annotation subdocument text stream
        /// </summary>
        public Int32 ccpAtn;

        /// <summary>
        /// Length of endnote subdocument text stream
        /// </summary>
        public Int32 ccpEdn;

        /// <summary>
        /// Length of textbox subdocument text stream
        /// </summary>
        public Int32 ccpTxbx;

        /// <summary>
        /// Length of header textbox subdocument text stream
        /// </summary>
        public Int32 ccpHdrTxbx;

        /// <summary>
        /// When there was insufficient memory for Word to expand the plcfbte at saved time, 
        /// the plcfbte is written to the file in a linked list of 512-byte pieces starting with this pn
        /// </summary>
        public Int32 pnFbpChpFirst;

        /// <summary>
        /// The page number of the lowest numbered page in the document that records CHPX FKP information
        /// </summary>
        public Int32 pnChpFirst;

        /// <summary>
        /// Count of CHPX FKPS recorded in file. In non-complex files if the number of 
        /// entries is the plcfbteChpx is less than this, the plcfbteChpx is incomplete
        /// </summary>
        public Int32 cpnBteChp;

        /// <summary>
        /// When there was isufficient memory for Word to expand the plcfbte at save time, 
        /// the plcfbte is written to the file in a linked list of 512-byte pieces starting with this pn
        /// </summary>
        public Int32 pnFbpPapFirst;

        /// <summary>
        /// The page number of the lowest numbered page in the document that records PAPX FKP information
        /// </summary>
        public Int32 pnPapFirst;

        /// <summary>
        /// Count of PAPX FKPS recorded in file.<br/>
        /// In non-complex files if the number of entries in the plcfbtePapx is 
        /// less than this, the plcfbtePapx is incomplete.
        /// </summary>
        public Int32 cpnBtePap;


        public Int32 pnFbpLvcFirst;

        public Int32 pnLvcFirst;

        public Int32 cpnBteLvc;

        public Int32 fcIslandFirst;

        public Int32 fcIslandLim;

        public UInt16 Cfclcb;

        public Int32 fcFtshfOrig;

        public UInt32 lcbStshfOrig;

        public Int32 fcStshf;

        public UInt32 lcbStshf;

        public Int32 fcPlcffndRef;

        public UInt32 lcbPlcffndRef;

        public Int32 fcPlcffndTxt;

        public UInt32 lcbPlcffndTxt;

        public Int32 fcPlcfandRef;

        public UInt32 lcbPlcfandRef;

        public Int32 fcPlcfandTxt;

        public UInt32 lcbPlcfandTxt;

        public Int32 fcPlcfSed;

        public UInt32 lcbPlcfSed;

        public Int32 fcPlcfphe;

        public UInt32 lcbPlcfphe;

        public Int32 fcSttbfglsy;

        public UInt32 lcbSttbfglsy;

        public Int32 fcPlcfglsy;

        public UInt32 lcbPlcfglsy;

        public Int32 fcPlcfhdd;

        public UInt32 lcbPlcfhdd;

        public Int32 fcPlcfbteChpx;

        public UInt32 lcbPlcfbteChpx;

        public Int32 fcPlcfbtePapx;

        public UInt32 lcbPlcfbtePapx;

        public Int32 fcPlcfsea;

        public UInt32 lcbPlcfsea;

        public Int32 fcSttbfffn;

        public UInt32 lcbSttbfffn;

        public Int32 fcPlcffldMom;

        public UInt32 lcbPlcffdldMom;

        public Int32 fcPlcffldHdr;

        public UInt32 lcbPlcffldHdr;

        public Int32 fcPlcffldFtn;

        public UInt32 lcbPlcffldFtn;

        public Int32 fcPlcffldAtn;

        public UInt32 lcbPlcffldAtn;

        public Int32 fcPlcffldMcr;

        public UInt32 lcbPlcffldMcr;

        public Int32 fcSttbfbkmk;

        public UInt32 lcbSttbfbkmk;

        public Int32 fcPlcfbkf;

        public UInt32 lcbPlcfbkf;

        public Int32 fcPlcfbkl;

        public UInt32 lcbPlcfbkl;

        public Int32 fcCmds;

        public UInt32 lcbCmds;

        public Int32 fcplcmcr;

        public UInt32 lcbPlcmcr;

        public Int32 fcSttbfmcr;

        public UInt32 lcbSttbfmcr;

        public Int32 fcPrDrvr;

        public UInt32 lcbPrDrvr;

        public Int32 fcPrEnvPort;

        public UInt32 lcbPrEnvPort;

        public Int32 fcPrEnvLand;

        public UInt32 lcbPrEnvLand;

        public Int32 fcWss;

        public UInt32 lcbWss;

        public Int32 fcDop;

        public UInt32 lcbDop;

        public Int32 fcSttbfAssoc;

        public UInt32 lcbSttbfAssoc;

        public Int32 fcClx;

        public UInt32 lcbClx;

        public Int32 fcPlcfpgdFtn;

        public Int32 fcAutosaveSource;

        public UInt32 lcbAutosaveSource;

        public Int32 fcGrpXstAtnOwners;

        public UInt32 lcbGrpXstAtnOwners;

        public Int32 fcSttbfAtnbkmk;

        public UInt32 lcbSttbfAtnbkmk;

        public Int32 fcPlcdoaMom;

        public UInt32 lcbPlcdiaMom;

        public Int32 fcPlcdoaHdr;

        public UInt32 lcbPlcdoaHdr;

        public Int32 fcPlcspaMom;

        public UInt32 lcbPlcspaMom;

        public Int32 fcPlcspaHdr;

        public UInt32 lcbPlcspaHdr;

        public Int32 fcPlcfAtnbkf;
        public UInt32 lcbPlcfAtnbkf;

        public Int32 fcPlcfAtnbkl;
        public UInt32 lcbPlcfAtnbkl;

        //Page 152

        public Int32 fcPms;
        public UInt32 lcbPms;

        public Int32 fcFormFldSttbs;
        public UInt32 lcbFormFldSttbs;

        public Int32 fcPlcfendRef;
        public UInt32 lcbPlcfendRef;

        public Int32 fcPlcfendTxt;
        public UInt32 lcbPlcfendTxt;

        public Int32 fcPlcffldEdn;
        public UInt32 lcbPlcffldEdn;

        public Int32 fcPlcfpgEdn;
        public UInt32 lcbPlcfpgEdn;

        public Int32 fcDggInfo;
        public UInt32 lcbDggInfo;

        public Int32 fcSttbfRMark;
        public UInt32 lcbSttbfRMark;

        //Page 153

        public Int32 fcSttbCaption;
        public UInt32 lcbSttbCaption;

        public Int32 fcSttbAutoCaption;
        public UInt32 lcbSttbAutoCaption;

        public Int32 fcPlcfwkb;
        public UInt32 lcbPlcfwkb;

        public Int32 fcPlcfspl;
        public UInt32 lcbPlcfspl;

        public Int32 fcPlcftxbxTxt;
        public UInt32 lcbPlcftxbxTxt;

        public Int32 fcPlcffldTxt;
        public UInt32 lcbPlcffldTxt;

        public Int32 fcPlcfhdrtxbxTxt;
        public UInt32 lcbPlcfhdrtxbxTxt;

        public Int32 fcPlcfldHdrTxbx;
        public UInt32 lcbPlcfldHdrTxbx;

        public Int32 fcStwUser;

        //Page 154

        public UInt32 lcbStwUser;

        public Int32 fcSttbttmbd;
        public UInt32 lcbSttbttmbd;

        public Int32 fcCookieData;
        public UInt32 lcbCookieData;

        public Int32 fcSttbfIntlFld;
        public UInt32 lcbSttbfIntlFld;

        public Int32 fcRouteSlip;
        public UInt32 lcbRouteSlip;

        public Int32 fcSttbSavedBy;
        public UInt32 lcbSttbSavedBy;

        public Int32 fcSttbFnm;
        public UInt32 lcbSttbFnm;

        // Page 155

        public Int32 fcPlcfLst;
        public UInt32 lcbPlcfLst;

        public Int32 fcPlfLfo;
        public UInt32 lcbPlfLfo;

        public Int32 fcPlcftxbxBkd;
        public UInt32 lcbPlcftxbxBkd;

        public Int32 fcPlcftxbxHdrBkd;
        public UInt32 lcbPlcftxbxHdrBkd;

        public Int32 fcDocUndoWord9;
        public UInt32 lcbDocUndoWord9;

        public Int32 fcRgbuse;
        public UInt32 lcbRgbuse;

        public Int32 fcUsp;
        public UInt32 lcbUsp;

        public Int32 fcUskf;
        public UInt32 lcbUskf;

        public Int32 fcPlcupcRgbuse;
        public UInt32 lcbPlcupcRgbuse;

        public Int32 fcPlcupcUsp;
        public UInt32 lcbPlcupcUsp;

        public Int32 fcSttbGlsyStyle;
        public UInt32 lcbSttbGlsyStyle;

        public Int32 fcPlgosl;

        // Page 156

        public UInt32 lcbPlgosl;

        public Int32 fcPlcocx;
        public UInt32 lcbPlcocx;

        public Int32 fcPlcfbteLvc;
        public UInt32 lcbPlcfbteLvc;

        public UInt32 dwLowDateTime;
        public UInt32 dwHighDateTime;

        public Int32 fcPlcflvcPre10;
        public UInt32 lcbPlcflvcPre10;

        public Int32 fcPlcasumy;
        public UInt32 lcbPlcasumy;

        public Int32 fcPlcfgram;
        public UInt32 lcbPlcfgram;

        public Int32 fcSttbListNames;
        public UInt32 lcbSttbListNames;

        public Int32 fcSttbfUssr;
        public UInt32 lcbSttbfUssr;

        public Int32 fcPlcfTch;
        public UInt32 lcbPlcfTch;

        public Int32 fcRmdfThreading;
        public UInt32 lcbRmdfThreading;

        public Int32 fcMid;
        
        //Page 157

        public UInt32 lcbMid;

        public Int32 fcSttbRgtplc;
        public UInt32 lcbSttbRgtplc;

        public Int32 fcMsoEnvelope;
        public UInt32 lcbMsoEnvelope;

        public Int32 fcPlcflad;
        public UInt32 lcbPlcflad;

        public Int32 fcRgdofr;
        public UInt32 lcbRgdofr;

        public Int32 fcPlcosl;
        public UInt32 lcbPlcosl;

        public Int32 fcPlcfcookieOld;
        public UInt32 lcbPlcfcookieOld;

        //Page 158

        public Int32 fcUnused;
        public UInt32 lcbUnused;

        public Int32 fcPlcfpgp;
        public UInt32 lcbPlcfpgp;

        public Int32 fcPlcfuim;
        public UInt32 lcbPlcfuim;

        public Int32 fcPlfguidUim;
        public UInt32 lcbPlfguidUim;

        public Int32 fcAtrdExtra;
        public UInt32 lcbAtrdExtra;

        public Int32 fcPlrsid;
        public UInt32 lcbPlrsid;

        public Int32 fcSttbfBkmkFactoid;
        public UInt32 lcbSttbfBkmkFactoid;

        //Page 159

        public Int32 fcPlcfBkfFactoid;
        public UInt32 lcbPlcfBkfFactoid;

        public Int32 fcPlcfcookie;
        public UInt32 lcbPlcfcookie;

        public Int32 fcPlcfBklFactoid;
        public UInt32 lcbPlcfBklFactoid;

        public Int32 fcFactoidData;
        public UInt32 lcbFactoidData;

        public Int32 fcDocUndo;
        public UInt32 lcbDocUndo;

        public Int32 fcSttbfBkmkFcc;
        public UInt32 lcbSttbfBkmkFcc;

        //Page 160

        public Int32 fcPlcfBkfFcc;
        public UInt32 lcbPlcfBkfFcc;

        public Int32 fcPlcfBklFcc;
        public UInt32 lcbPlcfBklFcc;

        public Int32 fcSttbfbkmkBPRepairs;
        public UInt32 lcbSttbfbkmkBPRepairs;

        public Int32 fcPlcfbkfBPRepairs;
        public UInt32 lcbPlcfbkfBPRepairs;

        public Int32 fcPlcfbklBPRepairs;
        public UInt32 lcbPlcfbklBPRepairs;

        public Int32 fcPmsNew;

        //Page 161

        public UInt32 lcbPmsNew;

        public Int32 fcODSO;
        public UInt32 lcbODSO;

        public Int32 fcPlcfpmiOldXP;
        public UInt32 lcbPlcfpmiOldXP;

        public Int32 fcPlcfpmiNewXP;
        public UInt32 lcbPlcfpmiNewXP;

        public Int32 fcPlcfpmiMixedXP;
        public UInt32 lcbPlcfpmiMixedXP;

        public Int32 fcEncryptedProps;
        public UInt32 lcbEncryptedProps;

        public Int32 fcPlcffactoid;
        public UInt32 lcbPlcffactoid;

        public Int32 fcPlcflvcOldXp;

        //Page 162

        public UInt32 lcbPlcflvcOldXp;

        public Int32 fcPlcflvcNewXp;
        public UInt32 lcbPlcflvcNewXp;

        public Int32 fcPlcflvcMixedXp;
        public UInt32 lcbPlcflvcMixedXp;

        public Int32 fcHplxsdr;
        public UInt32 lcbHplxsdr;

        public Int32 fcSttbfBkmSdt;
        public UInt32 lcbSttbfBkmSdt;

        public Int32 fcPlcfBkfSdt;
        public UInt32 lcbPlcfBkfSdt;

        //Page 163

        public Int32 fcPlcfBklSdt;
        public UInt32 lcbPlcfBklSdt;

        public Int32 fcCustomXForm;
        public UInt32 lcbCustomXForm;

        public Int32 fcSttbfBkmkProt;
        public UInt32 lcbSttbfBkmkProt;

        public Int32 fcPlcfBkfProt;
        public UInt32 lcbPlcfBkfProt;

        public Int32 fcPlcfBklProt;
        public UInt32 lcbPlcfBklProt;

        public Int32 fcSttbProtUser;
        public UInt32 lcbSttbProtUser;

        //Page 164

        public Int32 fcPlcftpc;
        public UInt32 lcbPlcftpc;

        public Int32 fcPlcfpmiOld;
        public UInt32 lcbPlcfpmiOld;

        public Int32 fcPlcfpmiOldInline;
        public UInt32 lcbPlcfpmiOldInline;

        public Int32 fcPlcfpmiNew;
        public UInt32 lcbPlcfpmiNew;

        public Int32 fcPlcfpmiNewInline;
        public UInt32 lcbPlcfpmiNewInline;

        public Int32 fcPlcflvcOld;
        public UInt32 lcbPlcflvcOld;

        public Int32 fcPlcflvcOldInline;
        public UInt32 lcbPlcflvcOldInline;

        public Int32 fcPlcflvcNew;
        public UInt32 lcbPlcflvcNew;

        public Int32 fcPlcflvcNewInline;
        public UInt32 lcbPlcflvcNewInline;

        //Page 165

        public Int32 fcAfd;
        public UInt32 lcbAfd;
        public UInt16 cswNew;
        public UInt16 nFib;
        public UInt16 cQuickSavesNew;

        //*****************************************************************************************
        //                                                                              CONSTRUCTOR
        //*****************************************************************************************

        /// <summary>
        /// Parses the File Information Block of a Compound Document
        /// </summary>
        /// <param name="wordDocument">The "WordDocument" stream</param>
        public FileInformationBlock(VirtualStream st)
        {
            try
            {
                //read the first 1472 bytes (FIB)
                byte[] bytes = new byte[1472];
                st.Read(bytes, 0, 1472, 0);

                //start parsing the variables
                wIdent = System.BitConverter.ToUInt16(bytes, 0);
                nFib = System.BitConverter.ToUInt16(bytes, 2);
                nProduct = System.BitConverter.ToUInt16(bytes, 4);
                Lid = System.BitConverter.ToUInt16(bytes, 6);
                pnNext = System.BitConverter.ToInt16(bytes, 8);
                fDot = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0001);
                fGlsy = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0002);
                fComplex = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0002);
                fHasPic = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0008);
                cQuickSaves = (UInt16)(((int)System.BitConverter.ToUInt16(bytes, 10) & 0x00F0) >> 4);
                fEncrypted = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0100);
                fWhichTblStm = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0200);
                fReadOnlyRecommended = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0400);
                fWriteReservation = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x0800);
                fExtChar = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x1000);
                fLoadOverwrite = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x2000);
                fFarEast = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x4000);
                fCrypto = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 10), 0x8000);
                nFibBack = System.BitConverter.ToUInt16(bytes, 12);
                lKey = System.BitConverter.ToInt32(bytes, 14);
                Envr = bytes[18];
                fMac = Utils.BitmaskToBool((int)bytes[19], 0x01);
                fEmptySpecial = Utils.BitmaskToBool((int)bytes[19], 0x02);
                fLoadOverridePage = Utils.BitmaskToBool((int)bytes[19], 0x04);
                fFutureSavedUndo = Utils.BitmaskToBool((int)bytes[19], 0x08);
                fWord97Saved = Utils.BitmaskToBool((int)bytes[19], 0x10);
                Chs = System.BitConverter.ToUInt16(bytes, 20);
                chsTables = System.BitConverter.ToUInt16(bytes, 22);
                fcMin = System.BitConverter.ToInt32(bytes, 24);
                fcMac = System.BitConverter.ToInt32(bytes, 28);
                Csw = System.BitConverter.ToUInt16(bytes, 32);
                wMagicCreated = System.BitConverter.ToInt16(bytes, 34);
                wMagicRevised = System.BitConverter.ToInt16(bytes, 36);
                wMagicCreatedPrivate = System.BitConverter.ToInt16(bytes, 38);
                wMagicRevisedPrivate = System.BitConverter.ToInt16(bytes, 40);
                lidFE = System.BitConverter.ToInt16(bytes, 60);
                Clw = System.BitConverter.ToUInt16(bytes, 62);
                cbMac = System.BitConverter.ToInt32(bytes, 64);
                lProductCreated = System.BitConverter.ToInt32(bytes, 68);
                lProductRevised = System.BitConverter.ToInt32(bytes, 72);
                ccpText = System.BitConverter.ToInt32(bytes, 76);
                ccpFtn = System.BitConverter.ToInt32(bytes, 80);
                ccpHdr = System.BitConverter.ToInt32(bytes, 84);
                ccpMcr = System.BitConverter.ToInt32(bytes, 88);
                ccpAtn = System.BitConverter.ToInt32(bytes, 92);
                ccpEdn = System.BitConverter.ToInt32(bytes, 96);
                ccpTxbx = System.BitConverter.ToInt32(bytes, 100);
                ccpHdrTxbx = System.BitConverter.ToInt32(bytes, 104);
                pnFbpChpFirst = System.BitConverter.ToInt32(bytes, 108);
                pnChpFirst = System.BitConverter.ToInt32(bytes, 112);
                cpnBteChp = System.BitConverter.ToInt32(bytes, 116);
                pnFbpPapFirst = System.BitConverter.ToInt32(bytes, 120);
                pnPapFirst = System.BitConverter.ToInt32(bytes, 124);
                cpnBtePap = System.BitConverter.ToInt32(bytes, 128);
                pnFbpLvcFirst = System.BitConverter.ToInt32(bytes, 132);
                pnLvcFirst = System.BitConverter.ToInt32(bytes, 136);
                cpnBteLvc = System.BitConverter.ToInt32(bytes, 140);
                fcIslandFirst = System.BitConverter.ToInt32(bytes, 144);
                fcIslandLim = System.BitConverter.ToInt32(bytes, 148);
                Cfclcb = System.BitConverter.ToUInt16(bytes, 152);
                fcFtshfOrig = System.BitConverter.ToInt32(bytes, 154);
                lcbStshfOrig = System.BitConverter.ToUInt32(bytes, 158);
                fcStshf = System.BitConverter.ToInt32(bytes, 162);
                lcbStshf = System.BitConverter.ToUInt32(bytes, 166);
                fcPlcffndRef = System.BitConverter.ToInt32(bytes, 170);
                lcbPlcffndRef = System.BitConverter.ToUInt32(bytes, 174);
                fcPlcffndTxt = System.BitConverter.ToInt32(bytes, 178);
                lcbPlcffndTxt = System.BitConverter.ToUInt32(bytes, 182);
                fcPlcfandRef = System.BitConverter.ToInt32(bytes, 186);
                lcbPlcfandRef = System.BitConverter.ToUInt32(bytes, 190);
                fcPlcfandTxt = System.BitConverter.ToInt32(bytes, 194);
                lcbPlcfandTxt = System.BitConverter.ToUInt32(bytes, 198);
                fcPlcfSed = System.BitConverter.ToInt32(bytes, 202);
                lcbPlcfSed = System.BitConverter.ToUInt32(bytes, 206);
                fcPlcfphe = System.BitConverter.ToInt32(bytes, 218);
                lcbPlcfphe = System.BitConverter.ToUInt32(bytes, 222);
                fcSttbfglsy = System.BitConverter.ToInt32(bytes, 226);
                lcbSttbfglsy = System.BitConverter.ToUInt32(bytes, 230);
                fcPlcfglsy = System.BitConverter.ToInt32(bytes, 234);
                lcbPlcfglsy = System.BitConverter.ToUInt32(bytes, 238);
                fcPlcfhdd = System.BitConverter.ToInt32(bytes, 242);
                lcbPlcfhdd = System.BitConverter.ToUInt32(bytes, 246);
                fcPlcfbteChpx = System.BitConverter.ToInt32(bytes, 250);
                lcbPlcfbteChpx = System.BitConverter.ToUInt32(bytes, 254);
                fcPlcfbtePapx = System.BitConverter.ToInt32(bytes, 258);
                lcbPlcfbtePapx = System.BitConverter.ToUInt32(bytes, 262);
                fcPlcfsea = System.BitConverter.ToInt32(bytes, 266);
                lcbPlcfsea = System.BitConverter.ToUInt32(bytes, 270);
                fcSttbfffn = System.BitConverter.ToInt32(bytes, 274);
                lcbSttbfffn = System.BitConverter.ToUInt32(bytes, 278);
                fcPlcffldMom = System.BitConverter.ToInt32(bytes, 282);
                lcbPlcffdldMom = System.BitConverter.ToUInt32(bytes, 286);
                fcPlcffldHdr = System.BitConverter.ToInt32(bytes, 290);
                lcbPlcffldHdr = System.BitConverter.ToUInt32(bytes, 294);
                fcPlcffldFtn = System.BitConverter.ToInt32(bytes, 298);
                lcbPlcffldFtn = System.BitConverter.ToUInt32(bytes, 302);
                fcPlcffldAtn = System.BitConverter.ToInt32(bytes, 306);
                lcbPlcffldAtn = System.BitConverter.ToUInt32(bytes, 310);
                fcPlcffldMcr = System.BitConverter.ToInt32(bytes, 314);
                lcbPlcffldMcr = System.BitConverter.ToUInt32(bytes, 318);
                fcSttbfbkmk = System.BitConverter.ToInt32(bytes, 322);
                lcbSttbfbkmk = System.BitConverter.ToUInt32(bytes, 326);
                fcPlcfbkf = System.BitConverter.ToInt32(bytes, 330);
                lcbPlcfbkf = System.BitConverter.ToUInt32(bytes, 334);
                fcPlcfbkl = System.BitConverter.ToInt32(bytes, 338);
                lcbPlcfbkl = System.BitConverter.ToUInt32(bytes, 342);
                fcCmds = System.BitConverter.ToInt32(bytes, 346);
                lcbCmds = System.BitConverter.ToUInt32(bytes, 350);
                fcplcmcr = System.BitConverter.ToInt32(bytes, 354);
                lcbPlcmcr = System.BitConverter.ToUInt32(bytes, 358);
                fcSttbfmcr = System.BitConverter.ToInt32(bytes, 362);
                lcbSttbfmcr = System.BitConverter.ToUInt32(bytes, 366);
                fcPrDrvr = System.BitConverter.ToInt32(bytes, 370);
                lcbPrDrvr = System.BitConverter.ToUInt32(bytes, 374);
                fcPrEnvPort = System.BitConverter.ToInt32(bytes, 378);
                lcbPrEnvPort = System.BitConverter.ToUInt32(bytes, 382);
                fcPrEnvLand = System.BitConverter.ToInt32(bytes, 386);
                lcbPrEnvLand = System.BitConverter.ToUInt32(bytes, 390);
                fcWss = System.BitConverter.ToInt32(bytes, 394);
                lcbWss = System.BitConverter.ToUInt32(bytes, 398);
                fcDop = System.BitConverter.ToInt32(bytes, 402);
                lcbDop = System.BitConverter.ToUInt32(bytes, 406);
                fcSttbfAssoc = System.BitConverter.ToInt32(bytes, 410);
                lcbSttbfAssoc = System.BitConverter.ToUInt32(bytes, 414);
                fcClx = System.BitConverter.ToInt32(bytes, 418);
                lcbClx = System.BitConverter.ToUInt32(bytes, 422);
                fcPlcfpgdFtn = System.BitConverter.ToInt32(bytes, 426);
                fcAutosaveSource = System.BitConverter.ToInt32(bytes, 434);
                //Page 151
                lcbAutosaveSource = System.BitConverter.ToUInt32(bytes, 438);
                fcGrpXstAtnOwners = System.BitConverter.ToInt32(bytes, 442);
                lcbGrpXstAtnOwners = System.BitConverter.ToUInt32(bytes, 446);
                fcSttbfAtnbkmk = System.BitConverter.ToInt32(bytes, 450);
                lcbSttbfAtnbkmk = System.BitConverter.ToUInt32(bytes, 454);
                fcPlcdoaMom = System.BitConverter.ToInt32(bytes, 458);
                lcbPlcdiaMom = System.BitConverter.ToUInt32(bytes, 462);
                fcPlcdoaHdr = System.BitConverter.ToInt32(bytes, 466);
                lcbPlcdoaHdr = System.BitConverter.ToUInt32(bytes, 470);
                fcPlcspaMom = System.BitConverter.ToInt32(bytes, 474);
                lcbPlcspaMom = System.BitConverter.ToUInt32(bytes, 478);
                fcPlcspaHdr = System.BitConverter.ToInt32(bytes, 482);
                lcbPlcspaHdr = System.BitConverter.ToUInt32(bytes, 486);
                fcPlcfAtnbkf = System.BitConverter.ToInt32(bytes, 490);
                lcbPlcfAtnbkf = System.BitConverter.ToUInt32(bytes, 494);
                fcPlcfAtnbkl = System.BitConverter.ToInt32(bytes, 498);
                lcbPlcfAtnbkl = System.BitConverter.ToUInt32(bytes, 502);
                //Page 152
                fcPms = System.BitConverter.ToInt32(bytes, 506);
                lcbPms = System.BitConverter.ToUInt32(bytes, 510);
                fcFormFldSttbs = System.BitConverter.ToInt32(bytes, 514);
                lcbFormFldSttbs = System.BitConverter.ToUInt32(bytes, 518);
                fcPlcfendRef = System.BitConverter.ToInt32(bytes, 522);
                lcbPlcfendRef = System.BitConverter.ToUInt32(bytes, 526);
                fcPlcfendTxt = System.BitConverter.ToInt32(bytes, 530);
                lcbPlcfendTxt = System.BitConverter.ToUInt32(bytes, 534);
                fcPlcffldEdn = System.BitConverter.ToInt32(bytes, 538);
                lcbPlcffldEdn = System.BitConverter.ToUInt32(bytes, 542);
                fcPlcfpgEdn = System.BitConverter.ToInt32(bytes, 546);
                lcbPlcfpgEdn = System.BitConverter.ToUInt32(bytes, 550);
                fcDggInfo = System.BitConverter.ToInt32(bytes, 554);
                lcbDggInfo = System.BitConverter.ToUInt32(bytes, 558);
                fcSttbfRMark = System.BitConverter.ToInt32(bytes, 562);
                lcbSttbfRMark = System.BitConverter.ToUInt32(bytes, 566);
                //Page 153
                fcSttbCaption = System.BitConverter.ToInt32(bytes, 570);
                lcbSttbCaption = System.BitConverter.ToUInt32(bytes, 574);
                fcSttbAutoCaption = System.BitConverter.ToInt32(bytes, 578);
                lcbSttbAutoCaption = System.BitConverter.ToUInt32(bytes, 582);
                fcPlcfwkb = System.BitConverter.ToInt32(bytes, 586);
                lcbPlcfwkb = System.BitConverter.ToUInt32(bytes, 590);
                fcPlcfspl = System.BitConverter.ToInt32(bytes, 594);
                lcbPlcfspl = System.BitConverter.ToUInt32(bytes, 598);
                fcPlcftxbxTxt = System.BitConverter.ToInt32(bytes, 602);
                lcbPlcftxbxTxt = System.BitConverter.ToUInt32(bytes, 606);
                fcPlcffldTxt = System.BitConverter.ToInt32(bytes, 610);
                lcbPlcffldTxt = System.BitConverter.ToUInt32(bytes, 614);
                fcPlcfhdrtxbxTxt = System.BitConverter.ToInt32(bytes, 618);
                lcbPlcfhdrtxbxTxt = System.BitConverter.ToUInt32(bytes, 622);
                fcPlcfldHdrTxbx = System.BitConverter.ToInt32(bytes, 626);
                lcbPlcfldHdrTxbx = System.BitConverter.ToUInt32(bytes, 630);
                fcStwUser = System.BitConverter.ToInt32(bytes, 634);
                //Page 154
                lcbStwUser = System.BitConverter.ToUInt32(bytes, 638);
                fcSttbttmbd = System.BitConverter.ToInt32(bytes, 642);
                lcbSttbttmbd = System.BitConverter.ToUInt32(bytes, 646);
                fcCookieData = System.BitConverter.ToInt32(bytes, 650);
                lcbCookieData = System.BitConverter.ToUInt32(bytes, 654);
                fcSttbfIntlFld = System.BitConverter.ToInt32(bytes, 706);
                lcbSttbfIntlFld = System.BitConverter.ToUInt32(bytes, 710);
                fcRouteSlip = System.BitConverter.ToInt32(bytes, 714);
                lcbRouteSlip = System.BitConverter.ToUInt32(bytes, 718);
                fcSttbSavedBy = System.BitConverter.ToInt32(bytes, 722);
                lcbSttbSavedBy = System.BitConverter.ToUInt32(bytes, 726);
                fcSttbFnm = System.BitConverter.ToInt32(bytes, 730);
                lcbSttbFnm = System.BitConverter.ToUInt32(bytes, 734);
                //Page 155
                fcPlcfLst = System.BitConverter.ToInt32(bytes, 738);
                lcbPlcfLst = System.BitConverter.ToUInt32(bytes, 742);
                fcPlfLfo = System.BitConverter.ToInt32(bytes, 746);
                lcbPlfLfo = System.BitConverter.ToUInt32(bytes, 750);
                fcPlcftxbxBkd = System.BitConverter.ToInt32(bytes, 754);
                lcbPlcftxbxBkd = System.BitConverter.ToUInt32(bytes, 758);
                fcPlcftxbxHdrBkd = System.BitConverter.ToInt32(bytes, 762);
                lcbPlcftxbxHdrBkd = System.BitConverter.ToUInt32(bytes, 766);
                fcDocUndoWord9 = System.BitConverter.ToInt32(bytes, 770);
                lcbDocUndoWord9 = System.BitConverter.ToUInt32(bytes, 774);
                fcRgbuse = System.BitConverter.ToInt32(bytes, 778);
                lcbRgbuse = System.BitConverter.ToUInt32(bytes, 782);
                fcUsp = System.BitConverter.ToInt32(bytes, 786);
                lcbUsp = System.BitConverter.ToUInt32(bytes, 790);
                fcUskf = System.BitConverter.ToInt32(bytes, 794);
                lcbUskf = System.BitConverter.ToUInt32(bytes, 798);
                fcPlcupcRgbuse = System.BitConverter.ToInt32(bytes, 802);
                lcbPlcupcRgbuse = System.BitConverter.ToUInt32(bytes, 806);
                fcPlcupcUsp = System.BitConverter.ToInt32(bytes, 810);
                lcbPlcupcUsp = System.BitConverter.ToUInt32(bytes, 814);
                fcSttbGlsyStyle = System.BitConverter.ToInt32(bytes, 818);
                lcbSttbGlsyStyle = System.BitConverter.ToUInt32(bytes, 822);
                fcPlgosl = System.BitConverter.ToInt32(bytes, 826);
                //Page 156
                lcbPlgosl = System.BitConverter.ToUInt32(bytes, 830);
                fcPlcocx = System.BitConverter.ToInt32(bytes, 834);
                lcbPlcocx = System.BitConverter.ToUInt32(bytes, 838);
                fcPlcfbteLvc = System.BitConverter.ToInt32(bytes, 842);
                lcbPlcfbteLvc = System.BitConverter.ToUInt32(bytes, 846);
                dwLowDateTime = System.BitConverter.ToUInt32(bytes, 850);
                dwHighDateTime = System.BitConverter.ToUInt32(bytes, 854);
                fcPlcflvcPre10 = System.BitConverter.ToInt32(bytes, 858);
                lcbPlcflvcPre10 = System.BitConverter.ToUInt32(bytes, 862);
                fcPlcasumy = System.BitConverter.ToInt32(bytes, 866);
                lcbPlcasumy = System.BitConverter.ToUInt32(bytes, 870);
                fcPlcfgram = System.BitConverter.ToInt32(bytes, 874);
                lcbPlcfgram = System.BitConverter.ToUInt32(bytes, 878);
                fcSttbListNames = System.BitConverter.ToInt32(bytes, 882);
                lcbSttbListNames = System.BitConverter.ToUInt32(bytes, 886);
                fcSttbfUssr = System.BitConverter.ToInt32(bytes, 890);
                lcbSttbfUssr = System.BitConverter.ToUInt32(bytes, 894);
                fcPlcfTch = System.BitConverter.ToInt32(bytes, 898);
                lcbPlcfTch = System.BitConverter.ToUInt32(bytes, 902);
                fcRmdfThreading = System.BitConverter.ToInt32(bytes, 906);
                lcbRmdfThreading = System.BitConverter.ToUInt32(bytes, 910);
                fcMid = System.BitConverter.ToInt32(bytes, 914);
                //Page 157
                lcbMid = System.BitConverter.ToUInt32(bytes, 918);
                fcSttbRgtplc = System.BitConverter.ToInt32(bytes, 922);
                lcbSttbRgtplc = System.BitConverter.ToUInt32(bytes, 926);
                fcMsoEnvelope = System.BitConverter.ToInt32(bytes, 930);
                lcbMsoEnvelope = System.BitConverter.ToUInt32(bytes, 934);
                fcPlcflad = System.BitConverter.ToInt32(bytes, 938);
                lcbPlcflad = System.BitConverter.ToUInt32(bytes, 942);
                fcRgdofr = System.BitConverter.ToInt32(bytes, 946);
                lcbRgdofr = System.BitConverter.ToUInt32(bytes, 950);
                fcPlcosl = System.BitConverter.ToInt32(bytes, 954);
                lcbPlcosl = System.BitConverter.ToUInt32(bytes, 958);
                fcPlcfcookieOld = System.BitConverter.ToInt32(bytes, 962);
                lcbPlcfcookieOld = System.BitConverter.ToUInt32(bytes, 966);
                //Page 158
                fcUnused = System.BitConverter.ToInt32(bytes, 1018);
                lcbUnused = System.BitConverter.ToUInt32(bytes, 1022);
                fcPlcfpgp = System.BitConverter.ToInt32(bytes, 1026);
                lcbPlcfpgp = System.BitConverter.ToUInt32(bytes, 1030);
                fcPlcfuim = System.BitConverter.ToInt32(bytes, 1034);
                lcbPlcfuim = System.BitConverter.ToUInt32(bytes, 1038);
                fcPlfguidUim = System.BitConverter.ToInt32(bytes, 1042);
                lcbPlfguidUim = System.BitConverter.ToUInt32(bytes, 1046);
                fcAtrdExtra = System.BitConverter.ToInt32(bytes, 1050);
                lcbAtrdExtra = System.BitConverter.ToUInt32(bytes, 1054);
                fcPlrsid = System.BitConverter.ToInt32(bytes, 1058);
                lcbPlrsid = System.BitConverter.ToUInt32(bytes, 1062);
                fcSttbfBkmkFactoid = System.BitConverter.ToInt32(bytes, 1066);
                lcbSttbfBkmkFactoid = System.BitConverter.ToUInt32(bytes, 1070);
                //Page 159
                fcPlcfBkfFactoid = System.BitConverter.ToInt32(bytes, 1074);
                lcbPlcfBkfFactoid = System.BitConverter.ToUInt32(bytes, 1078);
                fcPlcfcookie = System.BitConverter.ToInt32(bytes, 1082);
                lcbPlcfcookie = System.BitConverter.ToUInt32(bytes, 1086);
                fcPlcfBklFactoid = System.BitConverter.ToInt32(bytes, 1090);
                lcbPlcfBklFactoid = System.BitConverter.ToUInt32(bytes, 1094);
                fcFactoidData = System.BitConverter.ToInt32(bytes, 1098);
                lcbFactoidData = System.BitConverter.ToUInt32(bytes, 1102);
                fcDocUndo = System.BitConverter.ToInt32(bytes, 1106);
                lcbDocUndo = System.BitConverter.ToUInt32(bytes, 1010);
                fcSttbfBkmkFcc = System.BitConverter.ToInt32(bytes, 1114);
                lcbSttbfBkmkFcc = System.BitConverter.ToUInt32(bytes, 1118);
                //Page 160
                fcPlcfBkfFcc = System.BitConverter.ToInt32(bytes, 1122);
                lcbPlcfBkfFcc = System.BitConverter.ToUInt32(bytes, 1126);
                fcPlcfBklFcc = System.BitConverter.ToInt32(bytes, 1130);
                lcbPlcfBklFcc = System.BitConverter.ToUInt32(bytes, 1134);
                fcSttbfbkmkBPRepairs = System.BitConverter.ToInt32(bytes, 1138);
                lcbSttbfbkmkBPRepairs = System.BitConverter.ToUInt32(bytes, 1142);
                fcPlcfbkfBPRepairs = System.BitConverter.ToInt32(bytes, 1146);
                lcbPlcfbkfBPRepairs = System.BitConverter.ToUInt32(bytes, 1150);
                fcPlcfbklBPRepairs = System.BitConverter.ToInt32(bytes, 1154);
                lcbPlcfbklBPRepairs = System.BitConverter.ToUInt32(bytes, 1158);
                fcPmsNew = System.BitConverter.ToInt32(bytes, 1162);
                //Page 161
                lcbPmsNew = System.BitConverter.ToUInt32(bytes, 1158);
                fcODSO = System.BitConverter.ToInt32(bytes, 1170);
                lcbODSO = System.BitConverter.ToUInt32(bytes, 1174);
                fcPlcfpmiOldXP = System.BitConverter.ToInt32(bytes, 1178);
                lcbPlcfpmiOldXP = System.BitConverter.ToUInt32(bytes, 1182);
                fcPlcfpmiNewXP = System.BitConverter.ToInt32(bytes, 1186);
                lcbPlcfpmiNewXP = System.BitConverter.ToUInt32(bytes, 1190);
                fcPlcfpmiMixedXP = System.BitConverter.ToInt32(bytes, 1194);
                lcbPlcfpmiMixedXP = System.BitConverter.ToUInt32(bytes, 1198);
                fcEncryptedProps = System.BitConverter.ToInt32(bytes, 1202);
                lcbEncryptedProps = System.BitConverter.ToUInt32(bytes, 1206);
                fcPlcffactoid = System.BitConverter.ToInt32(bytes, 1210);
                lcbPlcffactoid = System.BitConverter.ToUInt32(bytes, 1214);
                fcPlcflvcOldXp = System.BitConverter.ToInt32(bytes, 1218);
                //Page 162
                lcbPlcflvcOldXp = System.BitConverter.ToUInt32(bytes, 1222);
                fcPlcflvcNewXp = System.BitConverter.ToInt32(bytes, 1226);
                lcbPlcflvcNewXp = System.BitConverter.ToUInt32(bytes, 1230);
                fcPlcflvcMixedXp = System.BitConverter.ToInt32(bytes, 1234);
                lcbPlcflvcMixedXp = System.BitConverter.ToUInt32(bytes, 1238);
                fcHplxsdr = System.BitConverter.ToInt32(bytes, 1242);
                lcbHplxsdr = System.BitConverter.ToUInt32(bytes, 1246);
                fcSttbfBkmSdt = System.BitConverter.ToInt32(bytes, 1250);
                lcbSttbfBkmSdt = System.BitConverter.ToUInt32(bytes, 1254);
                fcPlcfBkfSdt = System.BitConverter.ToInt32(bytes, 1258);
                lcbPlcfBkfSdt = System.BitConverter.ToUInt32(bytes, 1262);
                //Page 163
                fcPlcfBklSdt = System.BitConverter.ToInt32(bytes, 1266);
                lcbPlcfBklSdt = System.BitConverter.ToUInt32(bytes, 1270);
                fcCustomXForm = System.BitConverter.ToInt32(bytes, 1274);
                lcbCustomXForm = System.BitConverter.ToUInt32(bytes, 1278);
                fcSttbfBkmkProt = System.BitConverter.ToInt32(bytes, 1282);
                lcbSttbfBkmkProt = System.BitConverter.ToUInt32(bytes, 1286);
                fcPlcfBkfProt = System.BitConverter.ToInt32(bytes, 1290);
                lcbPlcfBkfProt = System.BitConverter.ToUInt32(bytes, 1294);
                fcPlcfBklProt = System.BitConverter.ToInt32(bytes, 1298);
                lcbPlcfBklProt = System.BitConverter.ToUInt32(bytes, 1302);
                fcSttbProtUser = System.BitConverter.ToInt32(bytes, 1306);
                lcbSttbProtUser = System.BitConverter.ToUInt32(bytes, 1310);
                //Page 164
                fcPlcftpc = System.BitConverter.ToInt32(bytes, 1314);
                lcbPlcftpc = System.BitConverter.ToUInt32(bytes, 1318);
                fcPlcfpmiOld = System.BitConverter.ToInt32(bytes, 1322);
                lcbPlcfpmiOld = System.BitConverter.ToUInt32(bytes, 1326);
                fcPlcfpmiOldInline = System.BitConverter.ToInt32(bytes, 1330);
                lcbPlcfpmiOldInline = System.BitConverter.ToUInt32(bytes, 1334);
                fcPlcfpmiNew = System.BitConverter.ToInt32(bytes, 1338);
                lcbPlcfpmiNew = System.BitConverter.ToUInt32(bytes, 1342);
                fcPlcfpmiNewInline = System.BitConverter.ToInt32(bytes, 1346);
                lcbPlcfpmiNewInline = System.BitConverter.ToUInt32(bytes, 1350);
                fcPlcflvcOld = System.BitConverter.ToInt32(bytes, 1354);
                lcbPlcflvcOld = System.BitConverter.ToUInt32(bytes, 1358);
                fcPlcflvcOldInline = System.BitConverter.ToInt32(bytes, 1362);
                lcbPlcflvcOldInline = System.BitConverter.ToUInt32(bytes, 1366);
                fcPlcflvcNew = System.BitConverter.ToInt32(bytes, 1370);
                lcbPlcflvcNew = System.BitConverter.ToUInt32(bytes, 1374);
                fcPlcflvcNewInline = System.BitConverter.ToInt32(bytes, 1378);
                lcbPlcflvcNewInline = System.BitConverter.ToUInt32(bytes, 1382);
                //Page 165
                fcAfd = System.BitConverter.ToInt32(bytes, 1458);
                lcbAfd = System.BitConverter.ToUInt32(bytes, 162);
                cswNew = System.BitConverter.ToUInt16(bytes, 1466);
                //nFib = System.BitConverter.ToUInt16(bytes, 1468);
                cQuickSavesNew = System.BitConverter.ToUInt16(bytes, 1470);
            }
            catch(Exception)
            {
                throw new ByteParseException("FIB");
            }
        }
    }
}
