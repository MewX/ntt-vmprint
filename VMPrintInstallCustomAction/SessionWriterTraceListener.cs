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
/*〈モジュール名〉                  SessionLogWriterTraceListener           */
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

namespace VMPrintInstallCustomAction
{
    public class SessionLogWriterTraceListener : TextWriterTraceListener , IDisposable
    {

        protected MemoryStream listenerStream;
        protected Session installSession;
        private bool isDisposed;

        public SessionLogWriterTraceListener(Session session)
            : base()
        {
            this.listenerStream = new MemoryStream();
            this.Writer = new StreamWriter(this.listenerStream);
            this.installSession = session;
        }

        #region IDisposable impelementation

        /// <summary>
        /// Releases resources held by the listener -
        /// will not automatically flush and write
        /// trace data to the install session log -
        /// call CloseAndWriteLog() before disposing
        /// to ensure data is written
        /// </summary>
        public new void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dipose(bool disposing)
        {
            if (!this.isDisposed)
            {
                if (disposing)
                {
                    if (this.Writer != null)
                    {
                        this.Writer.Close();
                        this.Writer.Dispose();
                        this.Writer = null;
                    }
                    if (this.listenerStream != null)
                    {
                        this.listenerStream.Close();
                        this.listenerStream.Dispose();
                        this.listenerStream = null;
                    }
                    if (this.installSession != null)
                        this.installSession = null;
                }
                this.isDisposed = true;
            }

            base.Dispose(disposing);
        }

        #endregion

        /// <summary>
        /// Closes the listener and writes accumulated
        /// trace data to the install session's log (Session.Log)
        /// The listener should not be used after calling
        /// this method, and should be disposed of.
        /// </summary>
        public void CloseAndWriteLog()
        {
            if (this.listenerStream != null &&
                this.installSession != null)
            {
                this.Flush();
                if (this.listenerStream.Length > 0)
                {
                    listenerStream.Position = 0;
                    using (StreamReader listenerStreamReader = new StreamReader(this.listenerStream))
                    {
                        this.installSession.Log(listenerStreamReader.ReadToEnd());
                    }
                }
                this.Close();
                this.Dispose();
                this.installSession = null;
            }
        }
    }
}
