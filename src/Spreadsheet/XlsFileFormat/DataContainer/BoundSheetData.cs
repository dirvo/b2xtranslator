/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using System.Collections;
using System.Text;
using System.IO;
using System.Collections.ObjectModel;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// This class stores the data from every Boundsheet 
    /// </summary>
    public class BoundSheetData : IVisitable
    {
        /// <summary>
        /// List with the cellrecords from the boundsheet 
        /// </summary>
        public List<LABELSST> LABELSSTList;
        public List<MULRK> MULRKList;
        public List<NUMBER> NUMBERList;
        public List<RK> SINGLERKList;
        public List<BLANK> BLANKList;
        public List<MULBLANK> MULBLANKList;
        public List<FORMULA> FORMULAList;
        public List<ARRAY> ARRAYList;
        public BOUNDSHEET boundsheetRecord;


        public String worksheetName;
        public int worksheetId;
        public String worksheetRef;
        public SortedList<Int32, RowData> rowDataTable;
        public List<ColumnInfoData> colInfoDataTable;
        public List<SharedFormulaData> sharedFormulaDataTable;
        public FormulaCell latestFormula;

        public MERGECELLS MERGECELLSData;

        /// <summary>
        /// Ctor 
        /// </summary>
        public BoundSheetData()
        {
            this.LABELSSTList = new List<LABELSST>();
            this.MULRKList = new List<MULRK>();
            this.NUMBERList = new List<NUMBER>();
            this.SINGLERKList = new List<RK>();
            this.MULBLANKList = new List<MULBLANK>();
            this.BLANKList = new List<BLANK>();
            this.FORMULAList = new List<FORMULA>();
            this.ARRAYList = new List<ARRAY>();
            this.rowDataTable = new SortedList<int, RowData>();
            this.sharedFormulaDataTable = new List<SharedFormulaData>();
            this.colInfoDataTable = new List<ColumnInfoData>();
            boundsheetRecord = null;

        }

        /// <summary>
        /// Adds a labelsst element to the internal list 
        /// </summary>
        /// <param name="labelsstdata">A LABELSSTData element</param>
        public void addLabelSST(LABELSST labelsst)
        {
            this.LABELSSTList.Add(labelsst);
            RowData rowData = this.getSpecificRow(labelsst.rw);
            StringCell cell = new StringCell();
            cell.setValue(labelsst.isst);
            cell.Col = labelsst.col;
            cell.Row = labelsst.rw;
            cell.TemplateID = labelsst.ixfe;
            rowData.addCell(cell);
        }

        /// <summary>
        /// Adds a mulrk record element to the internal list 
        /// a mulrk record stores some integer or floatingpoint values 
        /// </summary>
        /// <param name="mulrk">The MULRK Record</param>
        public void addMULRK(MULRK mulrk)
        {
            this.MULRKList.Add(mulrk);
            RowData rowData = this.getSpecificRow(mulrk.rw);
            if (mulrk.ixfe.Count == mulrk.rknumber.Count)
            {
                for (int i = 0; i < mulrk.rknumber.Count; i++)
                {
                    NumberCell cell = new NumberCell();
                    cell.Col = mulrk.colFirst + i;
                    cell.Row = mulrk.rw;
                    cell.setValue(mulrk.rknumber[i]);
                    cell.TemplateID = mulrk.ixfe[i];
                    rowData.addCell(cell);
                }
            }
        }

        /// <summary>
        /// Adds a NUMBER Biffrecord to the internal list 
        /// additional the method adds the specific NUMBER Data to a data container 
        /// </summary>
        /// <param name="number">NUMBER Biffrecord</param>
        public void addNUMBER(NUMBER number)
        {
            this.NUMBERList.Add(number);
            RowData rowData = this.getSpecificRow(number.rw);
            NumberCell cell = new NumberCell();
            cell.setValue(number.num);
            cell.Col = number.col;
            cell.Row = number.rw;
            cell.TemplateID = number.ixfe;
            rowData.addCell(cell);
        }

        /// <summary>
        /// Adds a RK Biffrecord to the internal list 
        /// additional the method adds the specific RK Data to a data container 
        /// </summary>
        /// <param name="number">NUMBER Biffrecord</param>
        public void addRK(RK singlerk)
        {
            this.SINGLERKList.Add(singlerk);
            RowData rowData = this.getSpecificRow(singlerk.rw);
            NumberCell cell = new NumberCell();
            cell.setValue(singlerk.num);
            cell.Col = singlerk.col;
            cell.Row = singlerk.rw;
            cell.TemplateID = singlerk.ixfe;
            rowData.addCell(cell);
        }


        /// <summary>
        /// Adds a BLANK Biffrecord to the internal list 
        /// additional the method adds the specific BLANK Data to a data container 
        /// </summary>
        /// <param name="number">NUMBER Biffrecord</param>
        public void addBLANK(BLANK blank)
        {
            this.BLANKList.Add(blank);
            RowData rowData = this.getSpecificRow(blank.rw);
            BlankCell cell = new BlankCell();

            cell.Col = blank.col;
            cell.Row = blank.rw;
            cell.TemplateID = blank.ixfe;
            rowData.addCell(cell);
        }


        /// <summary>
        /// Adds a mulblank record element to the internal list 
        /// a mulblank record stores some blank cells and their formating 
        /// </summary>
        /// <param name="mulrk">The MULRK Record</param>
        public void addMULBLANK(MULBLANK mulblank)
        {
            this.MULBLANKList.Add(mulblank);
            RowData rowData = this.getSpecificRow(mulblank.rw);

            for (int i = 0; i < mulblank.ixfe.Count; i++)
            {
                BlankCell cell = new BlankCell();
                cell.Col = mulblank.colFirst + i;
                cell.Row = mulblank.rw;
                cell.TemplateID = mulblank.ixfe[i];
                rowData.addCell(cell);
            }

        }

        /// <summary>
        /// Adds a formula BIFF RECORD to the formula list and 
        /// creates a new cell 
        /// </summary>
        /// <param name="formula"></param>
        public void addFORMULA(FORMULA formula)
        {
            this.FORMULAList.Add(formula);
            RowData rowData = this.getSpecificRow(formula.rw);
            FormulaCell cell = new FormulaCell();


            cell.setValue(formula.ptgStack);
            cell.Col = formula.col;
            cell.Row = formula.rw;
            cell.TemplateID = formula.ixfe;

            if (formula.fShrFmla)
            {
                ((FormulaCell)cell).isSharedFormula = true;
            }
            cell.calculatedValue = formula.calculatedValue;

            if (formula.boolValueSet)
            {
                cell.calculatedValue = formula.boolValue;
            }
            else if (formula.errorValue != 0)
            {
                cell.calculatedValue = formula.errorValue;
            }
            this.latestFormula = cell;
            rowData.addCell(cell);
        }

        /// <summary>
        /// Add a stringvalue to the formula 
        /// </summary>
        /// <param name="formulaString"></param>
        public void addFormulaString(string formulaString)
        {
            this.latestFormula.calculatedValue = formulaString;
        }

        /// <summary>
        /// Adds an ARRAY BIFF Record to the arraylist 
        /// </summary>
        /// <param name="array"></param>
        public void addARRAY(ARRAY array)
        {
            this.ARRAYList.Add(array);
        }

        /// <summary>
        /// Looks for a specific row number and if it doesn't exist the method will create the one.
        /// </summary>
        /// <param name="rownumber">the specific rownumber as integer</param>
        /// <returns></returns>
        public RowData getSpecificRow(int rownumber)
        {
            RowData rowData;
            if (this.rowDataTable.ContainsKey(rownumber))
            {
                rowData = (RowData)this.rowDataTable[rownumber];
            }
            else
            {
                rowData = new RowData(rownumber);
                this.rowDataTable.Add(rownumber, rowData);
            }
            return rowData;
        }

        /// <summary>
        /// Returns the stack at the position of the given values 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public Stack<AbstractPtg> getArrayData(UInt16 rw, UInt16 col)
        {
            Stack<AbstractPtg> stack = new Stack<AbstractPtg>();
            foreach (ARRAY array in this.ARRAYList)
            {
                if (array.colFirst == col && array.rwFirst == rw)
                {
                    stack = array.ptgStack;
                }
            }
            return stack;
        }

        /// <summary>
        /// Searches a cell at the specific Position 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public AbstractCellData getCellAtPosition(UInt16 rw, UInt16 col)
        {
            RowData rd = this.getSpecificRow((int)rw);
            AbstractCellData scell = null;
            foreach (AbstractCellData cell in rd.Cells)
            {
                if (cell.Row == rw && cell.Col == col)
                    scell = cell;
            }
            return scell;
        }

        /// <summary>
        /// Sets the value from a Formula Cell 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        public void setFormulaUsesArray(UInt16 rw, UInt16 col)
        {
            AbstractCellData cell = this.getCellAtPosition(rw, col);
            if (cell is FormulaCell)
            {
                ((FormulaCell)cell).usesArrayRecord = true;
            }
        }

        /// <summary>
        /// Add a shared formula to the internal list
        /// </summary>
        /// <param name="shrfmla"></param>
        public void addSharedFormula(SHRFMLA shrfmla)
        {
            SharedFormulaData sfd = new SharedFormulaData();
            sfd.colFirst = shrfmla.colFirst;
            sfd.colLast = shrfmla.colLast;
            sfd.rwFirst = shrfmla.rwFirst;
            sfd.rwLast = shrfmla.rwLast;
            sfd.setValue(shrfmla.ptgStack);
            int ID = this.sharedFormulaDataTable.Count;
            sfd.ID = ID;
            this.sharedFormulaDataTable.Add(sfd);
        }


        /// <summary>
        /// Checks if the formula cell with this coordinates is in the shared formula range 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns>Null if the cell isn't in a SharedFormula range
        ///          The SharedFormulaData Objekt if the cell is in the range</returns>
        public SharedFormulaData checkFormulaIsInShared(int rw, int col)
        {
            foreach (SharedFormulaData var in this.sharedFormulaDataTable)
            {
                if (var.checkFormulaIsInShared(rw, col))
                    return var;
            }
            return null;
        }

        /// <summary>
        /// Add the rowdata to the rowdataobject 
        /// </summary>
        /// <param name="row">ROW Biff Record</param>
        public void addRowData(ROW row)
        {
            RowData rowData = this.getSpecificRow(row.rw);

            rowData.height = new TwipsValue(row.miyRw);
            rowData.hidden = row.fDyZero;
            rowData.outlineLevel = row.iOutLevel;
            rowData.collapsed = row.fCollapsed;

            rowData.customFormat = row.fGhostDirty;
            rowData.style = row.ixfe_val;

            rowData.thickBot = row.fExDes;
            rowData.thickTop = row.fExAsc;

            rowData.maxSpan = row.colMac;
            rowData.minSpan = row.colMic;
        }

        /// <summary>
        /// Add the colinfo to the data object model 
        /// </summary>
        /// <param name="colinfo">COLINFO BIFF Record</param>
        public void addColData(COLINFO colinfo)
        {
            ColumnInfoData colinfoData = new ColumnInfoData();
            colinfoData.min = colinfo.colFirst;
            colinfoData.max = colinfo.colLast;

            colinfoData.widht = colinfo.coldx;
            colinfoData.customWidth = colinfo.fUserSet;

            colinfoData.hidden = colinfo.fHidden;
            colinfoData.bestFit = colinfo.fBestFit;
            colinfoData.phonetic = colinfo.fPhonetic;
            colinfoData.outlineLevel = colinfo.iOutLevel;
            colinfoData.collapsed = colinfo.fCollapsed;
            colinfoData.style = colinfo.ixfe;

            this.colInfoDataTable.Add(colinfoData); 
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<BoundSheetData>)mapping).Apply(this);
        }

        #endregion
    }
}
