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
/*〈モジュール名〉                  NativeMethods                           */
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
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace VMPrint
{
    internal static class NativeMethods
    {
        /*
        This code was adapted from Matthew Ephraim's Ghostscript.Net project -
        external dll definitions moved into NativeMethods to
        satisfy FxCop requirements
        https://github.com/mephraim/ghostscriptsharp
        */

        // TODO: replacing this
        #region Hooks into Ghostscript DLL
        [DllImport("gsdll64.dll", EntryPoint = "gsapi_new_instance")]
        internal static extern int CreateAPIInstance(out IntPtr pinstance, IntPtr caller_handle);

        [DllImport("gsdll64.dll", EntryPoint = "gsapi_init_with_args")]
        internal static extern int InitAPI(IntPtr instance, int argc, string[] argv);

        [DllImport("gsdll64.dll", EntryPoint = "gsapi_exit")]
        internal static extern int ExitAPI(IntPtr instance);

        [DllImport("gsdll64.dll", EntryPoint = "gsapi_delete_instance")]
        internal static extern void DeleteAPIInstance(IntPtr instance);
        #endregion

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool SetDllDirectory(string lpPathName);
    }
}
