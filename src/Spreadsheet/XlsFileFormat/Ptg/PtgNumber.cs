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
 * (INCLUDING NEGLIGE6NCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */
using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public enum PtgNumber : ushort
    {
        PtgExp = 0x01, 
        PtgTbl = 0x02, 
        PtgAdd = 0x03, 
        PtgSub = 0x04,
        PtgMul = 0x05,
        PtgDiv = 0x06, 
        PtgPower = 0x07,
        PtgConcat = 0x08, 
        PtgLt = 0x09, 
        PtgLe = 0x0A, 
        PtgEq = 0x0B,
        PtgGe = 0x0C, 
        PtgGt = 0x0D, 
        PtgNe = 0x0E, 
        PtgIsect = 0x0F, 
        PtgUnion = 0x10,
        PtgRange = 0x11, 
        PtgUplus = 0x12, 
        PtgUminus = 0x13, 
        PtgPercent = 0x14, 
        PtgParen = 0x15, 
        PtgMissArg = 0x16, 
        PtgStr = 0x17,
        Ptg0x18Sub = 0x18,
        Ptg0x19Sub = 0x19,
        PtgErr = 0x1C, 
        PtgBool = 0x1D, 
        PtgInt = 0x1E, 
        PtgNum = 0x1F, 
        PtgArray = 0x20, 
        PtgFunc = 0x21, 
        PtgFuncVar = 0x22, 
        PtgName = 0x23,
        PtgRef = 0x24, 
        PtgArea = 0x25, 
        PtgMemArea = 0x26, 
        PtgMemErr = 0x27, 
        PtgMemNoMem = 0x28, 
        PtgMemFunc = 0x29, 
        PtgRefErr = 0x2A, 
        PtgAreaErr = 0x2B, 
        PtgRefN = 0x2C, 
        PtgAreaN = 0x2D, 
        PtgNameX = 0x39, 
        PtgRef3d = 0x3A, 
        PtgArea3d = 0x3B, 
        PtgRefErr3d = 0x3C, 
        PtgAreaErr3d = 0x3D, /*
        PtgArray = 0x40, 
        PtgFunc = 0x41, 
        PtgFuncVar = 0x42, 
        PtgName = 0x43, 
        PtgRef = 0x44, 
        PtgArea = 0x45, 
        PtgMemArea = 0x46, 
        PtgMemErr = 0x47, 
        PtgMemNoMem = 0x48, 
        PtgMemFunc = 0x49, 
        PtgRefErr = 0x4A, 
        PtgAreaErr = 0x4B, 
        PtgRefN = 0x4C, 
        PtgAreaN = 0x4D,
        PtgNameX = 0x59, 
        PtgRef3d = 0x5A, 
        PtgArea3d = 0x5B,
        PtgRefErr3d = 0x5C,
        PtgAreaErr3d = 0x5D,
        PtgArray = 0x60, 
        PtgFunc = 0x61, 
        PtgFuncVar = 0x62,
        PtgName = 0x63,
        PtgRef = 0x64,
        PtgArea = 0x65, 
        PtgMemArea = 0x66,
        PtgMemErr = 0x67,
        PtgMemNoMem = 0x68,
        PtgMemFunc = 0x69,
        PtgRefErr = 0x6A, 
        PtgAreaErr = 0x6B, 
        PtgRefN = 0x6C, 
        PtgAreaN = 0x6D,
        PtgNameX = 0x79, 
        PtgRef3d = 0x7A,
        PtgArea3d = 0x7B, 
        PtgRefErr3d = 0x7C, 
        PtgAreaErr3d = 0x7D */ 
    }

    public enum Ptg0x18Sub : ushort
    {
        PtgElfLel = 0x01,
        PtgElfRw = 0x02, 
        PtgElfCol = 0x03, 
        PtgElfRwV = 0x06, 
        PtgElfColV = 0x07,
        PtgElfRadical = 0x0A, 
        PtgElfRadicalS = 0x0B,
        PtgElfColS = 0x0D,
        PtgElfColSV = 0x0F, 
        PtgElfRadicalLel = 0x10, 
        PtgSxName = 0x1D
    }

    public enum Ptg0x19Sub : ushort
    {
        PtgNotDocumented = 0x00,
        PtgAttrSemi = 0x01, 
        PtgAttrIf = 0x02,
        PtgAttrChoose = 0x04, 
        PtgAttrGoto = 0x08, 
        PtgAttrSum = 0x10,
        PtgAttrBaxcel1 = 0x20,
        PtgAttrBaxcel2 = 0x21, 
        PtgAttrSpace = 0x40, 
        PtgAttrSpaceSemi = 0x41
    }

    public enum PtgType : ushort
    {

        Operand,
        Operator
    }
}
