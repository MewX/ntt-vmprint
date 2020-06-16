/****************************************************************************/
/*                                                                          */
/*  Copyright (C) 2019                                                      */
/*          （株）ＮＴＴデータ                                              */
/*           第三金融事業本部 戦略ビジネス本部　システム企画担当            */
/*                                                                          */
/*  収容物  ＣＯＮＴＩＭＩＸＥ    VirtualPrinterDriver                      */
/*                                                                          */
/****************************************************************************/
/*--------------------------------------------------------------------------*/
/*〈対象業務名〉                                                            */
/*〈対象業務ＩＤ〉                                                          */
/*〈モジュール名〉                  GhostScript64                           */
/*〈モジュールＩＤ〉                                                        */
/*〈モジュール通番〉                                                        */
/*--------------------------------------------------------------------------*/
/* ＜適応ＯＳ＞                     Windows 10 XXX                          */
/* ＜開発環境＞                     Microsoft Visual Studio 2017            */
/*--------------------------------------------------------------------------*/
/* ＜開発システム名＞               ＣＯＮＴＩＭＩＸＥ                      */
/* ＜開発システム番号＞                                                     */
/*--------------------------------------------------------------------------*/
/* ＜開発担当名＞                   システム企画担当                        */
/* ＜電話番号＞                     050-5546-2418                           */
/*--------------------------------------------------------------------------*/
/* ＜設計者名＞                     範                                      */
/* ＜設計年月日＞                   2019年02月19日　　　　　                */
/* ＜設計修正者名＞                                                         */
/* ＜設計修正年月日及び修正ＩＤ＞                                           */
/*--------------------------------------------------------------------------*/
/* ＜ソース作成者名＞               範                                      */
/* ＜ソース作成年月日＞             2019年02月19日　　　　　                */
/* ＜ソース修正者名＞                                                       */
/* ＜ソース修正年月日及び修正ＩＤ＞                                         */
/*--------------------------------------------------------------------------*/
using System;
using System.Runtime.InteropServices;

namespace VMPrint
{
    internal class GhostScript64
    {
        /*
        This code was adapted from Matthew Ephraim's Ghostscript.Net project
        external dll definitions moved into NativeMethods to
        satisfy FxCop requirements
        https://github.com/mephraim/ghostscriptsharp
        */

        /// <summary>
        /// Calls the Ghostscript API with a collection of arguments to be passed to it
        /// </summary>
        public static void CallAPI(string[] args)
        {
            // Get a pointer to an instance of the Ghostscript API and run the API with the current arguments
            IntPtr gsInstancePtr;
            lock (resourceLock)
            {
                NativeMethods.CreateAPIInstance(out gsInstancePtr, IntPtr.Zero);
                try
                {
                    int result = NativeMethods.InitAPI(gsInstancePtr, args.Length, args);

                    if (result < 0)
                    {
                        throw new ExternalException("Ghostscript conversion error", result);
                    }
                }
                finally
                {
                    Cleanup(gsInstancePtr);
                }
            }
        }

        /// <summary>
        /// Frees up the memory used for the API arguments and clears the Ghostscript API instance
        /// </summary>
        private static void Cleanup(IntPtr gsInstancePtr)
        {
            NativeMethods.ExitAPI(gsInstancePtr);
            NativeMethods.DeleteAPIInstance(gsInstancePtr);
        }


        /// <summary>
        /// GS can only support a single instance, so we need to bottleneck any multi-threaded systems.
        /// </summary>
        private static object resourceLock = new object();
    }
}
