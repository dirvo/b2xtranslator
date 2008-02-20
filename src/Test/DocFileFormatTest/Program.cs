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
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.IO;
using System.Collections;

namespace DIaLOGIKa.b2xtranslator.DocFileFormatTest
{
    class Program
    {
        private static StorageReader reader;
        private static VirtualStream wordDocumentStream, tableStream;
        private static FileInformationBlock fib;
        private static string file, method;

        static void Main(string[] args)
        {
            try
            {
                //parse arguments
                parseArgs(args);

                reader = new StorageReader(file);

                //get the "WordDocument" stream
                wordDocumentStream = reader.GetStream("WordDocument");

                //parse the FIB
                fib = new FileInformationBlock(wordDocumentStream);

                //get the tablestream
                if (fib.fWhichTblStm)
                    tableStream = reader.GetStream("1Table");
                else
                    tableStream = reader.GetStream("0Table");

                //starting
                if (!fib.fComplex)
                {
                    if (method == "FKPPAPX")
                    {
                        testFKPPAPX();
                    }
                    else if (method == "FKPCHPX")
                    {
                        testFKPCHPX();
                    }
                    else if (method == "STSH")
                    {
                        testSTSH();
                    }
                    else if (method == "BIT")
                    {
                        testBIT();
                    }
                    else
                    {
                        printUsage();
                    }

                    Console.WriteLine("\nPress key to exit ...");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine(file + " has benn fast-saved. This format is currently not supported.");
                }

                reader.Close();
            }
            catch (ArgumentException ae)
            {
                printUsage();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                if(reader != null)
                reader.Close();
            }
        }

        private static void testBIT()
        {
            //10110
            BitArray bits = new BitArray(new bool[] {true, false, true, true, false});
            Console.WriteLine(Utils.BitArrayToUInt32(bits));
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
                FileInfo fi = new FileInfo(file);
                method = args[1];
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
            Console.WriteLine("Usage: Test {filename} {method}\n"+
                "methods can be:\n"+
                "FKPPAPX: prints the formatted disk pages width paragraph properties\n"+
                "FKPCHPX: prints the formatted disk pages width character properties\n"+
                "STSH: prints the contents of the stylesheet");
        }

        /// <summary>
        /// prints the contents of the stylesheet
        /// </summary>
        private static void testSTSH()
        {
            StyleSheet stsh = new StyleSheet(fib, tableStream);
            Console.WriteLine("Stylesheet contains " + stsh.Styles.Count + " Styles");

            for (int i=0; i<stsh.Styles.Count; i++)
            {
                Console.WriteLine("Style " + i);
                StyleSheetDescription std = stsh.Styles[i];
                if (std != null)
                {
                    Console.WriteLine("\tIdentifier: " + std.sti);
                    Console.WriteLine("\tStyle Kind: " + std.stk);
                    Console.WriteLine("\tBased On: " + std.istdBase);
                    Console.WriteLine("\tSemi Hidden: " + std.fSemiHidden); 
                    Console.WriteLine("\tName: " + std.xstzName);
                }
                else
                {
                    Console.WriteLine("\tEmpty Slot");
                }
            }
        }

        /// <summary>
        /// Method for testing FKP PAPX
        /// </summary>
        private static void testFKPPAPX()
        {
            //Get all PAPX FKPs
            List<FormattedDiskPagePAPX> papxFkps = FormattedDiskPagePAPX.GetAllPAPXFKPs(fib, wordDocumentStream, tableStream);
            Console.WriteLine("There are " + papxFkps.Count + " FKPs with PAPX in this file: \n");
            foreach (FormattedDiskPagePAPX fkp in papxFkps)
            {
                Console.Write("FKP matches on " + fkp.crun + " paragraphs: ");
                foreach (int mark in fkp.rgfc)
                {
                    Console.Write(mark + " ");
                }
                Console.WriteLine("");
                for (int i = 0; i < fkp.crun; i++)
                {
                    FormattedDiskPagePAPX.BX bx = fkp.rgbx[i];
                    ParagraphPropertyExceptions papx = fkp.grppapx[i];
                    Console.WriteLine("PAPX: has style " + papx.istd);
                    foreach (SinglePropertyModifier sprm in papx.grpprl)
                    {
                        Console.WriteLine(String.Format("\tSPRM: modifies " + sprm.Type + " property 0x{0:x4} (" + sprm.Arguments.Length + " bytes)", sprm.OpCode));
                    }
                }
                Console.WriteLine("");
            }
        }

        /// <summary>
        /// Method for testing FKP CHPX
        /// </summary>
        private static void testFKPCHPX()
        {
            List<FormattedDiskPageCHPX> chpxFkps = FormattedDiskPageCHPX.GetAllCHPXFKPs(fib, wordDocumentStream, tableStream);
            Console.WriteLine("There are " + chpxFkps.Count + " FKPs with CHPX in this file: \n");
            foreach (FormattedDiskPageCHPX fkp in chpxFkps)
            {
                Console.Write("FKP matches on " + fkp.crun + " characters: ");
                foreach (int mark in fkp.rgfc)
                {
                    Console.Write(mark + " ");
                }
                Console.WriteLine("");
                for (int i = 0; i < fkp.crun; i++)
                {
                    Console.WriteLine("CHPX:");
                    CharacterPropertyExceptions chpx = fkp.grpchpx[i];
                    foreach (SinglePropertyModifier sprm in chpx.grpprl)
                    {
                        Console.WriteLine(String.Format("\tSPRM: modifies " + sprm.Type + " property 0x{0:x4} (" + sprm.Arguments.Length + " bytes)", sprm.OpCode));
                    }
                }
            }
        }
    }
}
