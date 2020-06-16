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
/*〈モジュール名〉                  CustomAction                            */
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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

using VMPrintCore;

namespace VMPrintInstallCustomAction
{
    /// <summary>
    /// </summary>
    public class CustomActions
    {


        [CustomAction]
        public static ActionResult CheckIfPrinterNotInstalled(Session session)
        {
            ActionResult resultCode;
            SessionLogWriterTraceListener installTraceListener = new SessionLogWriterTraceListener(session);
            VMPrintInstaller installer = new VMPrintInstaller();
            installer.AddTraceListener(installTraceListener);
            try
            {
                if (installer.IsVMPrintPrinterInstalled())
                    resultCode = ActionResult.Success;
                else
                    resultCode = ActionResult.Failure;
            }
            finally
            {
                if (installTraceListener != null)
                    installTraceListener.Dispose();
            }

            return resultCode;
        }


        [CustomAction]
        public static ActionResult InstallVMPrintPrinter(Session session)
        {
            ActionResult printerInstalled;

            String driverSourceDirectory = session.CustomActionData["DriverSourceDirectory"];
            String outputCommand = session.CustomActionData["OutputCommand"];
            String outputCommandArguments = session.CustomActionData["OutputCommandArguments"];

            SessionLogWriterTraceListener installTraceListener = new SessionLogWriterTraceListener(session);
            installTraceListener.TraceOutputOptions = TraceOptions.DateTime;

            VMPrintInstaller installer = new VMPrintInstaller();
            installer.AddTraceListener(installTraceListener);
            try
            {


                if (installer.InstallVMPrintPrinter(driverSourceDirectory,
                                                      outputCommand,
                                                      outputCommandArguments))
                    printerInstalled = ActionResult.Success;
                else
                    printerInstalled = ActionResult.Failure;

                installTraceListener.CloseAndWriteLog();
            }
            finally
            {
                if (installTraceListener != null)
                    installTraceListener.Dispose();
                
            }
            return printerInstalled;
        }


        [CustomAction]
        public static ActionResult UninstallVMPrintPrinter(Session session)
        {
            ActionResult printerUninstalled;

            SessionLogWriterTraceListener installTraceListener = new SessionLogWriterTraceListener(session);
            installTraceListener.TraceOutputOptions = TraceOptions.DateTime;

            VMPrintInstaller installer = new VMPrintInstaller();
            installer.AddTraceListener(installTraceListener);
            try
            {
                if (installer.UninstallVMPrintPrinter())
                    printerUninstalled = ActionResult.Success;
                else
                    printerUninstalled = ActionResult.Failure;
                installTraceListener.CloseAndWriteLog();
            }
            finally
            {
                if (installTraceListener != null)
                    installTraceListener.Dispose();
            }
            return printerUninstalled;
        }
    }
}
