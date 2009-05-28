using System;
using System.IO;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// This class should extract the specific worksheet data. 
    /// </summary>
    public class WorksheetExtractor : Extractor
    {
        /// <summary>
        /// Datacontainer for the worksheet
        /// </summary>
        private WorkSheetData bsd;

        /// <summary>
        /// CTor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="bsd"> Boundsheetdata container</param>
        public WorksheetExtractor(VirtualStreamReader reader, WorkSheetData bsd)
            : base(reader) 
        {
            this.bsd = bsd;
            this.extractData(); 
        }

        /// <summary>
        /// Extracting the data from the stream 
        /// </summary>
        public override void extractData()
        {
            BiffHeader bh, latestbiff;
            BOF firstBOF = null;
            

            try
            {
                while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
                {
                    bh.id = (RecordType)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();

                    TraceLogger.DebugInternal("BIFF {0}\t{1}\t", bh.id, bh.length);
                    
                    if (bh.id == RecordType.EOF)
                    {
                        this.StreamReader.BaseStream.Seek(0, SeekOrigin.End);
                        
                    } 
                    else if (bh.id == RecordType.BOF)
                    {
                        BOF bof = new BOF(this.StreamReader, bh.id, bh.length);
                        
                        switch (bof.docType)
                        {
                            case BOF.DocumentType.WorkbookGlobals:
                            case BOF.DocumentType.Worksheet:
                                firstBOF = bof;
                                break;

                            case BOF.DocumentType.Chart:
                                // parse chart 

                                break;

                            default:
                                this.readUnkownFile();
                                break;
                        }
                    }
                    else if (bh.id == RecordType.LabelSst)
                    {
                        LabelSst labelsst = new LabelSst(this.StreamReader, bh.id, bh.length);
                        this.bsd.addLabelSST(labelsst); 
                    }
                    else if (bh.id == RecordType.MulRk) 
                    {
                        MulRk mulrk = new MulRk(this.StreamReader, bh.id, bh.length);
                        this.bsd.addMULRK(mulrk);
                    }
                    else if (bh.id == RecordType.Number)
                    {
                        Number number = new Number(this.StreamReader, bh.id, bh.length);
                        this.bsd.addNUMBER(number);
                    }
                    else if (bh.id == RecordType.RK)
                    {
                        RK rk = new RK(this.StreamReader, bh.id, bh.length);
                        this.bsd.addRK(rk);
                    }
                    else if (bh.id == RecordType.MergeCells)
                    {
                        MergeCells mergecells = new MergeCells(this.StreamReader, bh.id, bh.length);
                        this.bsd.MERGECELLSData = mergecells;
                    }
                    else if (bh.id == RecordType.Blank)
                    {
                        Blank blankcell = new Blank(this.StreamReader, bh.id, bh.length);
                        this.bsd.addBLANK(blankcell);
                    }
                    else if (bh.id == RecordType.MulBlank)
                    {
                        MulBlank mulblank = new MulBlank(this.StreamReader, bh.id, bh.length);
                        this.bsd.addMULBLANK(mulblank);
                    }
                    else if (bh.id == RecordType.Formula)
                    {
                        Formula formula = new Formula(this.StreamReader, bh.id, bh.length);
                        this.bsd.addFORMULA(formula);
                        TraceLogger.DebugInternal(formula.ToString());
                    }
                    else if (bh.id == RecordType.Array)
                    {
                        ARRAY array = new ARRAY(this.StreamReader, bh.id, bh.length);
                        this.bsd.addARRAY(array); 
                    }
                    else if (bh.id == RecordType.ShrFmla)
                    {
                        ShrFmla shrfmla = new ShrFmla(this.StreamReader, bh.id, bh.length);
                        this.bsd.addSharedFormula(shrfmla); 

                    }
                    else if (bh.id == RecordType.String)
                    {
                        STRING formulaString = new STRING(this.StreamReader, bh.id, bh.length);
                        this.bsd.addFormulaString(formulaString.value); 

                    }
                    else if (bh.id == RecordType.Row)
                    {
                        Row row = new Row(this.StreamReader, bh.id, bh.length);
                        this.bsd.addRowData(row); 

                    }
                    else if (bh.id == RecordType.ColInfo)
                    {
                        ColInfo colinfo = new ColInfo(this.StreamReader, bh.id, bh.length);
                        this.bsd.addColData(colinfo);
                    }
                    else if (bh.id == RecordType.DefColWidth)
                    {
                        DefColWidth defcolwidth = new DefColWidth(this.StreamReader, bh.id, bh.length);
                        this.bsd.addDefaultColWidth(defcolwidth.cchdefColWidth);
                    }
                    else if (bh.id == RecordType.DefaultRowHeight)
                    {
                        DefaultRowHeight defrowheigth = new DefaultRowHeight(this.StreamReader, bh.id, bh.length);
                        this.bsd.addDefaultRowData(defrowheigth); 
                    }
                    else if (bh.id == RecordType.LeftMargin)
                    {
                        LeftMargin leftm = new LeftMargin(this.StreamReader, bh.id, bh.length);
                        this.bsd.leftMargin = leftm.value; 
                    }
                    else if (bh.id == RecordType.RightMargin)
                    {
                        RightMargin rightm = new RightMargin(this.StreamReader, bh.id, bh.length);
                        this.bsd.rightMargin = rightm.value; 
                    }
                    else if (bh.id == RecordType.TopMargin)
                    {
                        TopMargin topm = new TopMargin(this.StreamReader, bh.id, bh.length);
                        this.bsd.topMargin = topm.value; 
                    }
                    else if (bh.id == RecordType.BottomMargin)
                    {
                        BottomMargin bottomm = new BottomMargin(this.StreamReader, bh.id, bh.length);
                        this.bsd.bottomMargin = bottomm.value; 
                    }
                    else if (bh.id == RecordType.Setup)
                    {
                        Setup setup = new Setup(this.StreamReader, bh.id, bh.length);
                        this.bsd.addSetupData(setup);
                    }
                    else if (bh.id == RecordType.HLink)
                    {
                        long oldStreamPos = this.StreamReader.BaseStream.Position; 
                        try
                        {

                            HLink hlink = new HLink(this.StreamReader, bh.id, bh.length);
                            bsd.addHyperLinkData(hlink); 
                        }
                        catch (Exception ex)
                        {
                            this.StreamReader.BaseStream.Seek(oldStreamPos, System.IO.SeekOrigin.Begin);
                            this.StreamReader.BaseStream.Seek(bh.length, System.IO.SeekOrigin.Current);
                            TraceLogger.Debug("Link parse error");
                            TraceLogger.Error(ex.StackTrace);
                        }
                    }

                    else
                    {
                        // this else statement is used to read BiffRecords which aren't implemented 
                        byte[] buffer = new byte[bh.length];
                        buffer = this.StreamReader.ReadBytes(bh.length);
                    }


                    latestbiff = bh; 
                }
            }
            catch (Exception ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Error(ex.StackTrace); 
                TraceLogger.Debug(ex.ToString());
            }
        }

        /// <summary>
        /// This method should read over every record which is inside a file in the worksheet file 
        /// For example this could be the diagram "file" 
        /// A diagram begins with the BOF Biffrecord and ends with the EOF record. 
        /// </summary>
        public void readUnkownFile(){
            BiffHeader bh;
            try
            {
                do
                {
                    bh.id = (RecordType)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();
                    this.StreamReader.ReadBytes(bh.length);
                } while (bh.id != RecordType.EOF); 
            }
            catch (Exception ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
        }
    }
}
