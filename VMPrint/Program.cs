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
/*〈モジュール名〉                  Program                                 */
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
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace VMPrint
{
    public class Program
    {
        #region Other constants
        const string traceSourceName = "VMPrint";
        #endregion

        static TraceSource logEventSource = new TraceSource(traceSourceName);

        [STAThread]
        static void Main(string[] args)
        {
            // Install the global exception handler
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(Application_UnhandledException);

            String standardInputFilename = Path.GetTempFileName();
            String outputFilename = String.Empty;
            try
            {
                if (!GetPdfOutputFilename(ref outputFilename))
                {
                    // TODO: remove me
                    MessageBox.Show("Unknown error when getting PDF output file name.", "Error",
                                     MessageBoxButtons.YesNo,
                                     MessageBoxIcon.Question);
                }

                using (BinaryReader standardInputReader = new BinaryReader(Console.OpenStandardInput()))
                {
                    using (FileStream standardInputFile = new FileStream(standardInputFilename, FileMode.Create, FileAccess.ReadWrite))
                    {
                        standardInputReader.BaseStream.CopyTo(standardInputFile);

                        //標準INPUTから読み出すデータがない場合、設定ダイアログを表示する。
                        if (standardInputFile.Length == 0)
                        {
                            SettingDlg dlg = new SettingDlg();
                            dlg.ShowDialog();
                            return;
                        }
                    }
                }

                // Only set absolute minimum parameters, let the postscript input
                // dictate as much as possible
                XpsPdfPrinter.ConvertXpsToBitmapToPdf(standardInputFilename, outputFilename);

                // Display PDF file.
                DisplayPdf(outputFilename);
            }
            catch (IOException ioEx)
            {
                // We couldn't delete, or create a file
                // because it was in use
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          Properties.Resources.ERROR_COULD_NOT_WRITE +
                                          Environment.NewLine +
                                          Properties.Resources.EXCEPTION_MESSAGE_PREFIX + ioEx.Message);
                DisplayErrorMessage(Properties.Resources.ERROR_DIALOG_CAPTION,
                                    Properties.Resources.ERROR_COULD_NOT_WRITE + Environment.NewLine +
                                    String.Format(Properties.Resources.ERROR_IN_USE, outputFilename));
            }
            catch (UnauthorizedAccessException unauthorizedEx)
            {
                // Couldn't delete a file
                // because it was set to readonly
                // or couldn't create a file
                // because of permissions issues
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          Properties.Resources.ERROR_COULD_NOT_WRITE +
                                          Environment.NewLine +
                                          Properties.Resources.EXCEPTION_MESSAGE_PREFIX + unauthorizedEx.Message);
                DisplayErrorMessage(Properties.Resources.ERROR_DIALOG_CAPTION,
                                    Properties.Resources.ERROR_COULD_NOT_WRITE + Environment.NewLine +
                                    String.Format(Properties.Resources.ERROR_INSUFFICIENT_PRIVILEGES, outputFilename));
            }
            finally
            {
                try
                {
                    File.Delete(standardInputFilename);
                }
                catch
                {
                    logEventSource.TraceEvent(TraceEventType.Warning,
                                              (int)TraceEventType.Warning,
                                              String.Format(Properties.Resources.WARN_FILE_NOT_DELETED, standardInputFilename));
                }

                // More temp files to clean up.
                logEventSource.Flush();
            }
        }

        /// <summary>
        /// All unhandled exceptions will bubble their way up here -
        /// a final error dialog will be displayed before the crash and burn
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void Application_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            logEventSource.TraceEvent(TraceEventType.Critical,
                                      (int)TraceEventType.Critical,
                                      ((Exception)e.ExceptionObject).Message + Environment.NewLine +
                                                                        ((Exception)e.ExceptionObject).StackTrace);
            DisplayErrorMessage(Properties.Resources.ERROR_DIALOG_CAPTION,
                                Properties.Resources.ERROR_UNEXPECTED_ERROR);
        }

        static bool GetPdfOutputFilename(ref String outputFile)
        {
            bool filenameRetrieved = false;

            //Read property file
            String outputDir = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.OutputDir);
            String branchNo = Properties.Settings.Default.BranchNo;
            String machineNo = Properties.Settings.Default.MachineNo;
            long serialNo = Properties.Settings.Default.SerialNo;

            //PDF出力ディレクトリが設定ありの場合、自動的にPDFの名前を生成する。それ以外の場合、ダイアログを表示し、ユーザーに指定させる。
            if (String.IsNullOrEmpty(outputDir) == false && CreateDirectory(outputDir) == true)
            {
                serialNo += 1;
                outputFile = Path.Combine(outputDir, branchNo + machineNo + Convert.ToString(serialNo).PadLeft(10, '0') + ".pdf");
                Properties.Settings.Default.SerialNo = serialNo;
                Properties.Settings.Default.Save();
                filenameRetrieved = true;
            }
            else
            {
                using (SetOutputFilename dialogOwner = new SetOutputFilename())
                {
                    dialogOwner.TopMost = true;
                    dialogOwner.TopLevel = true;
                    dialogOwner.Show(); // Form won't actually show - Application.Run() never called
                                        // but having a topmost/toplevel owner lets us bring the SaveFileDialog to the front
                    dialogOwner.BringToFront();
                    using (SaveFileDialog pdfFilenameDialog = new SaveFileDialog())
                    {
                        pdfFilenameDialog.AddExtension = true;
                        pdfFilenameDialog.AutoUpgradeEnabled = true;
                        pdfFilenameDialog.CheckPathExists = true;
                        pdfFilenameDialog.Filter = Properties.Resources.FILENAME_FILTER;
                        pdfFilenameDialog.ShowHelp = false;
                        pdfFilenameDialog.Title = Properties.Resources.FILENAME_TITLE;
                        pdfFilenameDialog.ValidateNames = true;
                        //OKボタンをクリックする場合、選択されたPDFファイルの名前を戻る
                        if (pdfFilenameDialog.ShowDialog(dialogOwner) == DialogResult.OK)
                        {
                            outputFile = pdfFilenameDialog.FileName;
                            filenameRetrieved = true;
                        }
                    }
                    dialogOwner.Close();
                }
            }
            return filenameRetrieved;
        }

        static public bool CreateDirectory(String strDir)
        {
            bool createSuccess = true;

            if (Directory.Exists(strDir) == false)
            {
                try
                {
                    Directory.CreateDirectory(strDir);
                }
                catch (Exception ex)
                {
                    logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          ex.Message);
                    createSuccess = false;
                }
            }

            return createSuccess;
        }

        /// <summary>
        /// Opens the PDF in the default viewer
        /// if the OpenAfterPrint app setting is "True"
        /// and the file extension is .PDF
        /// </summary>
        /// <param name="pdfFilename"></param>
        static void DisplayPdf(String pdfFilename)
        {
            //自動的にPDFを開きがONの場合、PDFを作成した後、自動的に開く
            if (Properties.Settings.Default.OpenAfterPrint &&
                !String.IsNullOrEmpty(Path.GetExtension(pdfFilename)) &&
                (Path.GetExtension(pdfFilename).ToUpper() == ".PDF"))
            {
                Process.Start(pdfFilename);
            }
        }

        /// <summary>
        /// Displays up a topmost, OK-only message box for the error message
        /// </summary>
        /// <param name="boxCaption">The box's caption</param>
        /// <param name="boxMessage">The box's message</param>
        static void DisplayErrorMessage(String boxCaption,
                                        String boxMessage)
        {

            MessageBox.Show(boxMessage,
                            boxCaption,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);

        }
    }
}
