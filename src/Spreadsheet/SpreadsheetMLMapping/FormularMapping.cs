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
using System.Text;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.BiffRecords;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg; 
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer; 

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class FormularInfixMapping
    {


        public static String mapFormula(Object stack, ExcelContext xlsContext)
        {
            Stack<String> resultStack = new Stack<string>();
            try
            {

            
            String namexValue; 
            if (stack is Stack<AbstractPtg>)
            {
                Stack<AbstractPtg> opStack = new Stack<AbstractPtg>(((Stack<AbstractPtg>)stack).ToArray());
                if (opStack.Count == 0)
                    throw new Exception(); 

                while (opStack.Count != 0)
                {
                    AbstractPtg ptg = opStack.Pop();

                    if (ptg is PtgInt || ptg is PtgNum || ptg is PtgBool || ptg is PtgMissArg) 
                    {
                        resultStack.Push(ptg.getData());
                    }
                    else if (ptg is PtgRef)
                    {
                       PtgRef ptgref = (PtgRef)ptg;
                       resultStack.Push(ExcelHelperClass.intToABCString((int)ptgref.col, (ptgref.rw + 1).ToString(),ptgref.colRelative,ptgref.rwRelative));
                    }
                    else if (ptg is PtgUplus || ptg is PtgUminus)
                    {
                        String buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = ptg.getData() + resultStack.Pop();

                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgParen)
                    {
                        String buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = "(" + resultStack.Pop() + ")";

                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgPercent)
                    {
                        String buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = resultStack.Pop() + ptg.getData();

                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgAdd || ptg is PtgDiv || ptg is PtgMul ||
                                ptg is PtgSub || ptg is PtgPower || ptg is PtgGt ||
                                ptg is PtgGe || ptg is PtgLt || ptg is PtgLe ||
                                ptg is PtgEq || ptg is PtgNe || ptg is PtgConcat ||
                                ptg is PtgUnion || ptg is PtgIsect
                        )
                    {
                        String buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = ptg.getData() + resultStack.Pop();
                        buffer = resultStack.Pop() + buffer;
                        resultStack.Push(buffer);
                    } 
                    else if ( ptg is PtgStr) 
                    {
                        resultStack.Push("\"" + ptg.getData() + "\"");
                    }
                    else if (ptg is PtgArea)
                    {
                        String buffer = "";
                        PtgArea ptgref = (PtgArea)ptg;
                        buffer = ExcelHelperClass.intToABCString((int)ptgref.colFirst, (ptgref.rwFirst + 1).ToString(), ptgref.colFirstRelative, ptgref.rwFirstRelative);
                        buffer = buffer + ":" + ExcelHelperClass.intToABCString((int)ptgref.colLast, (ptgref.rwLast + 1).ToString(), ptgref.colLastRelative, ptgref.rwLastRelative);


                        resultStack.Push(buffer); 
                    }
                    else if (ptg is PtgAttrSum)
                    {
                        String buffer = "";
                        PtgAttrSum ptgref = (PtgAttrSum)ptg;
                        buffer = ptg.getData() + "(" + resultStack.Pop() + ")";
                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgAttrIf || ptg is PtgAttrGoto || ptg is PtgAttrSemi
                    || ptg is PtgAttrChoose || ptg is PtgAttrSpace)
                    {
                        
                    }
                    else if (ptg is PtgExp)
                    {
                        PtgExp ptgexp = (PtgExp) ptg;
                        Stack<AbstractPtg> newptgstack = xlsContext.CurrentSheet.getArrayData(ptgexp.rw, ptgexp.col);
                        resultStack.Push(FormularInfixMapping.mapFormula(newptgstack, xlsContext));
                        xlsContext.CurrentSheet.setFormulaUsesArray(ptgexp.rw, ptgexp.col); 

                    }
                    else if (ptg is PtgRef3d)
                    {
                        PtgRef3d ptgr3d = (PtgRef3d)ptg;
                        String refstring = xlsContext.XlsDoc.workBookData.getIXTIString(ptgr3d.ixti);
                        String cellref = ExcelHelperClass.intToABCString((int)ptgr3d.col, (ptgr3d.rw + 1).ToString(), ptgr3d.colRelative, ptgr3d.rwRelative);

                        resultStack.Push(refstring + "!" + cellref); 
                    }
                    else if (ptg is PtgArea3d)
                    {
                        PtgArea3d ptga3d = (PtgArea3d)ptg;
                        String refstring = xlsContext.XlsDoc.workBookData.getIXTIString(ptga3d.ixti);
                        String buffer = "";
                        buffer = ExcelHelperClass.intToABCString((int)ptga3d.colFirst, (ptga3d.rwFirst + 1).ToString(), ptga3d.colFirstRelative, ptga3d.rwFirstRelative);
                        buffer = buffer + ":" + ExcelHelperClass.intToABCString((int)ptga3d.colLast, (ptga3d.rwLast + 1).ToString(), ptga3d.colLastRelative, ptga3d.rwLastRelative);

                        resultStack.Push(refstring + "!" + buffer);
                    }
                    else if (ptg is PtgNameX)
                    {
                        PtgNameX ptgnx = (PtgNameX)ptg;
                        String opstring = xlsContext.XlsDoc.workBookData.getExternNameByRef(ptgnx.ixti, ptgnx.nameindex); 
                        namexValue = opstring;
                        resultStack.Push(opstring); 
                    }
                    else if (ptg is PtgFunc)
                    {
                        PtgFunc ptgf = (PtgFunc)ptg;
                        FtabValues value = (FtabValues)ptgf.tab;
                        String buffer = value.ToString();
                        buffer.Replace("_", ".");

                        // no param 
                        if (value == FtabValues.NA || value == FtabValues.PI ||
                            value == FtabValues.TRUE || value == FtabValues.FALSE
                            || value == FtabValues.RAND || value == FtabValues.NOW
                            || value == FtabValues.TODAY
                            )
                        {
                            buffer += "()";
                        }
                        // One param 
                        else if (value == FtabValues.ISNA || value == FtabValues.ISERROR ||
                                    value == FtabValues.SIN || value == FtabValues.COS ||
                            value == FtabValues.TAN || value == FtabValues.ATAN ||
                            value == FtabValues.SQRT || value == FtabValues.EXP ||
                            value == FtabValues.LN || value == FtabValues.LOG10 ||
                            value == FtabValues.ABS || value == FtabValues.INT ||
                            value == FtabValues.SIGN || value == FtabValues.LEN ||
                            value == FtabValues.VALUE || value == FtabValues.NOT
                            || value == FtabValues.DAY || value == FtabValues.MONTH
                            || value == FtabValues.YEAR || value == FtabValues.HOUR
                            || value == FtabValues.MINUTE || value == FtabValues.SECOND
                            || value == FtabValues.AREAS || value == FtabValues.ROWS
                            || value == FtabValues.COLUMNS || value == FtabValues.TRANSPOSE
                            || value == FtabValues.TYPE || value == FtabValues.ASIN
                            || value == FtabValues.ACOS || value == FtabValues.ISREF
                            || value == FtabValues.CHAR || value == FtabValues.LOWER
                            || value == FtabValues.UPPER || value == FtabValues.PROPER
                            || value == FtabValues.TRIM || value == FtabValues.CODE
                            || value == FtabValues.ISERR || value == FtabValues.ISTEXT
                            || value == FtabValues.ISNUMBER || value == FtabValues.ISBLANK
                            || value == FtabValues.T || value == FtabValues.N
                            || value == FtabValues.DATEVALUE || value == FtabValues.TIMEVALUE
                            || value == FtabValues.CLEAN || value == FtabValues.MDETERM
                            || value == FtabValues.MINVERSE || value == FtabValues.FACT
                            || value == FtabValues.ISNONTEXT || value == FtabValues.ISLOGICAL
                            || value == FtabValues.LENB || value == FtabValues.DBCS
                            || value == FtabValues.SINH || value == FtabValues.COSH
                            || value == FtabValues.TANH || value == FtabValues.ASINH
                            || value == FtabValues.ACOSH || value == FtabValues.ATANH
                            || value == FtabValues.INFO || value == FtabValues.ERROR_TYPE
                            || value == FtabValues.GAMMALN || value == FtabValues.EVEN
                            || value == FtabValues.FISHER || value == FtabValues.FISHERINV
                            || value == FtabValues.NORMSDIST || value == FtabValues.NORMSINV
                            || value == FtabValues.ODD || value == FtabValues.RADIANS
                            || value == FtabValues.DEGREES || value == FtabValues.COUNTBLANK
                            || value == FtabValues.DATESTRING || value == FtabValues.PHONETIC
                            )
                        {
                            buffer += "(" + resultStack.Pop() + ")";
                        }
                        // two params 
                        else if (value == FtabValues.ROUND || value == FtabValues.REPT ||
                                value == FtabValues.MOD || value == FtabValues.TEXT
                                || value == FtabValues.ATAN2 || value == FtabValues.EXACT
                                || value == FtabValues.MMULT || value == FtabValues.ROUNDUP
                                || value == FtabValues.ROUNDDOWN || value == FtabValues.FREQUENCY
                                || value == FtabValues.CHIDIST || value == FtabValues.CHIINV
                                || value == FtabValues.COMBIN || value == FtabValues.FLOOR
                                || value == FtabValues.CEILING || value == FtabValues.PERMUT
                                || value == FtabValues.SUMXMY2 || value == FtabValues.SUMX2MY2
                                || value == FtabValues.SUMX2PY2 || value == FtabValues.CHITEST
                                || value == FtabValues.CORREL || value == FtabValues.COVAR
                                || value == FtabValues.FTEST || value == FtabValues.INTERCEPT
                                || value == FtabValues.PEARSON || value == FtabValues.RSQ
                                || value == FtabValues.STEYX || value == FtabValues.SLOPE
                                || value == FtabValues.LARGE || value == FtabValues.SMALL
                                || value == FtabValues.QUARTILE || value == FtabValues.PERCENTILE
                                || value == FtabValues.TRIMMEAN || value == FtabValues.TINV
                                || value == FtabValues.POWER || value == FtabValues.COUNTIF
                            || value == FtabValues.NUMBERSTRING


                            )
                        {
                            buffer += "(";
                            String buffer2 = resultStack.Pop();
                            buffer2 = resultStack.Pop() + "," + buffer2;

                            buffer += buffer2 + ")";
                        }
                        // Three params 
                        else if (value == FtabValues.MID || value == FtabValues.DCOUNT ||
                                    value == FtabValues.DSUM || value == FtabValues.DAVERAGE ||
                                    value == FtabValues.DMIN || value == FtabValues.DMAX ||
                                    value == FtabValues.DSTDEV || value == FtabValues.DVAR
                                || value == FtabValues.MIRR || value == FtabValues.DATE
                                || value == FtabValues.TIME || value == FtabValues.SLN
                                || value == FtabValues.DPRODUCT || value == FtabValues.DSTDEVP
                                || value == FtabValues.DVARP || value == FtabValues.DCOUNTA
                                || value == FtabValues.MIDB || value == FtabValues.DGET
                                || value == FtabValues.CONFIDENCE || value == FtabValues.CRITBINOM
                                || value == FtabValues.EXPONDIST || value == FtabValues.FDIST

                                || value == FtabValues.FINV || value == FtabValues.GAMMAINV
                                || value == FtabValues.LOGNORMDIST || value == FtabValues.LOGINV
                                || value == FtabValues.NEGBINOMDIST || value == FtabValues.NORMINV
                                || value == FtabValues.STANDARDIZE || value == FtabValues.POISSON
                                || value == FtabValues.TDIST || value == FtabValues.FORECAST
                            )
                        {
                            buffer += "(";
                            String buffer2 = resultStack.Pop();
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer += buffer2 + ")";
                        }
                        // four params 
                        else if (value == FtabValues.REPLACE || value == FtabValues.SYD
                                || value == FtabValues.REPLACEB || value == FtabValues.BINOMDIST
                                || value == FtabValues.GAMMADIST || value == FtabValues.HYPGEOMDIST
                                || value == FtabValues.NORMDIST || value == FtabValues.WEIBULL
                                || value == FtabValues.TTEST || value == FtabValues.ISPMT

                            )
                        {
                            buffer += "(";
                            String buffer2 = resultStack.Pop();
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer += buffer2 + ")";
                        }
                        if ((int)value != 0xff)
                        {
                            
                        }
                        resultStack.Push(buffer);
                       
                    }
                    else if (ptg is PtgFuncVar)
                    {
                        PtgFuncVar ptgfv = (PtgFuncVar)ptg;
                        // is Ftab or Cetab 
                        if (!ptgfv.fCelFunc)
                        {
                            FtabValues value = (FtabValues)ptgfv.tab;
                            String buffer = value.ToString();
                            buffer.Replace("_", ".");
                            // 1 to x parameter
                            if (value == FtabValues.COUNT || value == FtabValues.IF ||
                                value == FtabValues.ISNA || value == FtabValues.ISERROR ||
                                value == FtabValues.AVERAGE || value == FtabValues.MAX ||
                                value == FtabValues.MIN || value == FtabValues.SUM ||
                                value == FtabValues.ROW || value == FtabValues.COLUMN ||
                                value == FtabValues.NPV || value == FtabValues.STDEV ||
                                value == FtabValues.DOLLAR || value == FtabValues.FIXED ||
                                value == FtabValues.SIN || value == FtabValues.COS ||
                                value == FtabValues.LOOKUP || value == FtabValues.INDEX ||
                                value == FtabValues.AND || value == FtabValues.OR ||
                                value == FtabValues.VAR || value == FtabValues.LINEST ||
                                value == FtabValues.TREND || value == FtabValues.LOGEST ||
                                value == FtabValues.GROWTH || value == FtabValues.PV ||
                                value == FtabValues.FV || value == FtabValues.NPER ||
                                value == FtabValues.PMT || value == FtabValues.RATE ||
                                value == FtabValues.IRR || value == FtabValues.MATCH ||
                                value == FtabValues.WEEKDAY || value == FtabValues.OFFSET
                                || value == FtabValues.ARGUMENT || value == FtabValues.SEARCH
                                || value == FtabValues.CHOOSE || value == FtabValues.HLOOKUP
                                || value == FtabValues.VLOOKUP || value == FtabValues.LOG
                                || value == FtabValues.LEFT || value == FtabValues.RIGHT
                                || value == FtabValues.SUBSTITUTE || value == FtabValues.FIND
                                || value == FtabValues.CELL || value == FtabValues.DDB
                                || value == FtabValues.INDIRECT || value == FtabValues.IPMT
                                || value == FtabValues.PPMT || value == FtabValues.COUNTA
                                || value == FtabValues.PRODUCT || value == FtabValues.STDEVP
                                || value == FtabValues.VARP || value == FtabValues.TRUNC
                                || value == FtabValues.USDOLLAR || value == FtabValues.FINDB
                                || value == FtabValues.SEARCHB || value == FtabValues.LEFTB
                                || value == FtabValues.RIGHTB || value == FtabValues.RANK
                                || value == FtabValues.ADDRESS || value == FtabValues.DAYS360
                                || value == FtabValues.VDB || value == FtabValues.MEDIAN
                                || value == FtabValues.PRODUCT || value == FtabValues.DB
                                || value == FtabValues.AVEDEV || value == FtabValues.BETADIST
                                || value == FtabValues.BETAINV || value == FtabValues.PROB
                                || value == FtabValues.DEVSQ || value == FtabValues.GEOMEAN
                                || value == FtabValues.HARMEAN || value == FtabValues.SUMSQ
                                || value == FtabValues.KURT || value == FtabValues.SKEW
                                || value == FtabValues.ZTEST || value == FtabValues.PERCENTRANK
                                || value == FtabValues.MODE || value == FtabValues.CONCATENATE
                                || value == FtabValues.SUBTOTAL || value == FtabValues.SUMIF
                                || value == FtabValues.ROMAN || value == FtabValues.GETPIVOTDATA
                                || value == FtabValues.HYPERLINK || value == FtabValues.AVERAGEA
                                || value == FtabValues.MAXA || value == FtabValues.MINA
                                || value == FtabValues.STDEVPA || value == FtabValues.VARPA
                                || value == FtabValues.STDEVA || value == FtabValues.VARA
                                )
                            {
                                buffer += "(";
                                String buffer2 = resultStack.Pop();
                                for (int i = 1; i < ptgfv.cparams; i++)
                                {
                                    buffer2 = resultStack.Pop() + "," + buffer2;
                                }
                                buffer += buffer2 + ")";
                                resultStack.Push(buffer);
                            }
                            else if ((int)value == 0xFF)
                            {
                                buffer = "(";
                                String buffer2 = resultStack.Pop();
                                for (int i = 1; i < ptgfv.cparams-1; i++)
                                {
                                    buffer2 = resultStack.Pop() + "," + buffer2;
                                }
                                buffer += buffer2 + ")";
                                // take the additional Operator from the Operandstack 
                                buffer = resultStack.Pop() + buffer; 
                                resultStack.Push(buffer);
                            }
                            
                        }
                    }
                }
            }
            else
            {
                resultStack.Push(""); 
            }
        }
        catch (Exception)
        {
            resultStack.Push(""); 
        }
            return resultStack.Pop(); 
        }
    }
}
