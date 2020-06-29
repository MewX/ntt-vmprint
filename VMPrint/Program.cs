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
using System.Runtime.InteropServices;
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
            String tempPdfFilename = String.Empty;
            String pdfTextFilename = String.Empty;
            String outputFilename = String.Empty;
            Boolean needPrint = false;
            try
            {
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
                tempPdfFilename = Path.GetTempFileName();
                MessageBox.Show("1: " + standardInputFilename, tempPdfFilename,
                                 MessageBoxButtons.YesNo,
                                 MessageBoxIcon.Question);

                XpsPdfPrinter.ConvertXpsToBitmapToPdf(standardInputFilename, tempPdfFilename + ".pdf");

                // TODO: remove this.
                MessageBox.Show("2: " + standardInputFilename, tempPdfFilename,
                                 MessageBoxButtons.YesNo,
                                 MessageBoxIcon.Question);

                String[] ghostScriptArguments = { "-q", "-dBATCH", "-dNOPAUSE", "-dSAFER",  "-sDEVICE=pdfwrite",
                                                String.Format("-sOutputFile={0}", tempPdfFilename), standardInputFilename };
                GhostScript64.CallAPI(ghostScriptArguments); // TODO: remove this ones

                //通常ﾌﾟﾘﾝﾀｰと検索文字列が設定ありの場合、印刷対象のドキュメントを検索し、検索文字列が含まない場合、通常ﾌﾟﾘﾝﾀｰで印刷する。
                if (String.IsNullOrEmpty(Properties.Settings.Default.RealPrinterName) == false &&
                        String.IsNullOrEmpty(Properties.Settings.Default.Style) == false)
                {
                    string[] keywords = Properties.Settings.Default.Style.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    pdfTextFilename = Path.GetTempFileName();
                    Pdf2Text(tempPdfFilename, pdfTextFilename);
                    needPrint = true;

                    using (StreamReader pdfTextStreamReader = new StreamReader(pdfTextFilename))
                    {
                        while (pdfTextStreamReader.EndOfStream == false)
                        {
                            string textLine = pdfTextStreamReader.ReadLine();
                            for (int i = 0; i < keywords.Length; i++)
                            {
                                if (String.IsNullOrEmpty(keywords[i]) == false && textLine.Contains(keywords[i]) == true)
                                {
                                    needPrint = false;
                                    break;
                                }
                            }
                            if (needPrint == false)
                            {
                                break;
                            }
                        }
                    }

                    //通常ﾌﾟﾘﾝﾀｰで印刷すると判明する場合、通常ﾌﾟﾘﾝﾀｰで印刷を実施する。
                    if (needPrint == true)
                    {
                        String[] printArguments = { "-q", "-dPrinted", "-dBATCH", "-dNOPAUSE", "-dNOSAFER", "-sDEVICE=mswinpr2", "-dNumCopies=1",
                                                 "-sOutputFile=%printer%" + Properties.Settings.Default.RealPrinterName, standardInputFilename };
                        GhostScript64.CallAPI(printArguments); // TODO: remove this
                    }
                }

                //PDFに出力すると判明する場合、PDFファイルを作成する。
                if (needPrint == false)
                {
                    if (GetPdfOutputFilename(ref outputFilename))
                    {
                        // Remove the existing PDF file if present
                        File.Delete(outputFilename);
                        //Copy file
                        System.IO.File.Copy(tempPdfFilename, outputFilename, true);
                        //Open pdf
                        DisplayPdf(outputFilename);
                    }
                }
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
            catch (ExternalException ghostscriptEx) // TODO: replace GhostScript error handling
            {
                // Ghostscript error
                logEventSource.TraceEvent(TraceEventType.Error,
                                          (int)TraceEventType.Error,
                                          String.Format(Properties.Resources.ERROR_GHOSTSCRIPT_CONVERSION, ghostscriptEx.ErrorCode.ToString()) +
                                          Environment.NewLine +
                                          Properties.Resources.EXCEPTION_MESSAGE_PREFIX + ghostscriptEx.Message);
                //If transferred to another virtual printer and cancelled, -100 error will occur
                if (ghostscriptEx.ErrorCode != -100)
                {
                    DisplayErrorMessage(Properties.Resources.ERROR_DIALOG_CAPTION,
                                    Properties.Resources.ERROR_PDF_GENERATION + Environment.NewLine +
                                    String.Format(Properties.Resources.ERROR_GHOSTSCRIPT_CONVERSION, ghostscriptEx.ErrorCode.ToString()));
                }
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

                try
                {
                    //一時PDF存在する場合、削除する
                    if (tempPdfFilename != String.Empty)
                    {
                        File.Delete(tempPdfFilename);
                    }
                }
                catch
                {
                    logEventSource.TraceEvent(TraceEventType.Warning,
                                              (int)TraceEventType.Warning,
                                              String.Format(Properties.Resources.WARN_FILE_NOT_DELETED, tempPdfFilename));
                }

                try
                {
                    //一時PDFのTEXTファイル存在する場合、削除する
                    if (pdfTextFilename != String.Empty)
                    {
                        File.Delete(pdfTextFilename);
                    }
                }
                catch
                {
                    logEventSource.TraceEvent(TraceEventType.Warning,
                                              (int)TraceEventType.Warning,
                                              String.Format(Properties.Resources.WARN_FILE_NOT_DELETED, pdfTextFilename));
                }

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

        /// <summary>
        /// Converts PDF to text file
        /// </summary>
        /// <param name="pdfFilename">Pdf file path</param>
        /// <param name="textFilename">Text file path</param>
        static void Pdf2Text(String pdfFilename,
                             String textFilename)
        {
            using (Process pdfToTextProcess = new Process())
            {
                pdfToTextProcess.StartInfo.FileName = "PDFTOTEXT.EXE";
                pdfToTextProcess.StartInfo.WorkingDirectory = Application.StartupPath;
                pdfToTextProcess.StartInfo.Arguments = " -enc UTF-8 " + pdfFilename + " " + textFilename;
                pdfToTextProcess.StartInfo.UseShellExecute = false;
                pdfToTextProcess.StartInfo.RedirectStandardInput = true;
                pdfToTextProcess.StartInfo.RedirectStandardOutput = true;
                pdfToTextProcess.StartInfo.RedirectStandardError = true;
                pdfToTextProcess.StartInfo.CreateNoWindow = true;
                pdfToTextProcess.Start();
                pdfToTextProcess.WaitForExit();
            }
        }
    }
}
