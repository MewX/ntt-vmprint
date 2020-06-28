/****************************************************************************/
/*                                                                          */
/*  Copyright (C) 2020                                                      */
/*          （株）ＮＴＴデータ                                                */
/*           第三金融事業本部 戦略ビジネス本部　システム企画担当                  */
/*                                                                          */
/*  収容物  ＣＯＮＴＩＭＩＸＥ    VirtualPrinterDriver                         */
/*                                                                          */
/****************************************************************************/
/*--------------------------------------------------------------------------*/
/*〈対象業務名〉                                                              */
/*〈対象業務ＩＤ〉                                                            */
/*〈モジュール名〉                  XpsPdfPrinter                             */
/*〈モジュールＩＤ〉                                                          */
/*〈モジュール通番〉                                                          */
/*--------------------------------------------------------------------------*/
/* ＜適応ＯＳ＞                     Windows 10 XXX                           */
/* ＜開発環境＞                     Microsoft Visual Studio 2017             */
/*--------------------------------------------------------------------------*/
/* ＜開発システム名＞               ＣＯＮＴＩＭＩＸＥ                          */
/* ＜開発システム番号＞                                                       */
/*--------------------------------------------------------------------------*/
/* ＜開発担当名＞                   N/A                                      */
/* ＜電話番号＞                     N/A                                      */
/*--------------------------------------------------------------------------*/
/* ＜設計者名＞                     MewX                                     */
/* ＜設計年月日＞                   2020年06月20日　　　　　                   */
/* ＜設計修正者名＞                                                          */
/* ＜設計修正年月日及び修正ＩＤ＞                                              */
/*--------------------------------------------------------------------------*/
/* ＜ソース作成者名＞               MewX                                      */
/* ＜ソース作成年月日＞             2020年06月20日　　　　　                    */
/* ＜ソース修正者名＞                                                         */
/* ＜ソース修正年月日及び修正ＩＤ＞                                            */
/*--------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VMPrint
{
    class XpsPdfPrinter
    {
        /// <summary>
        /// Calls the Ghostscript API with a collection of arguments to be passed to it
        /// </summary>
        public static void PrintPdf(string[] args)
        {
        }

        /// <summary>
        /// Clean up temporary files.
        /// </summary>
        private static void Cleanup(IntPtr gsInstancePtr)
        {
            // TODO
        }

        public void NativeMSPrintPdf(string pdfFileName)
        {
            // initialize PrintDocument object
            PrintDocument doc = new PrintDocument()
            {
                PrinterSettings = new PrinterSettings()
                {
                    // set the printer to 'Microsoft Print to PDF'
                    PrinterName = "Microsoft Print to PDF",

                    // tell the object this document will print to file
                    PrintToFile = true,

                    // set the filename to whatever you like (full path)
                    PrintFileName = pdfFileName,
                }
            };

            doc.Print();
        }
    }
}
