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
/* ＜設計者名＞                     Amos Xia                                 */
/* ＜設計年月日＞                   2020年06月20日　　　　　                   */
/* ＜設計修正者名＞                                                          */
/* ＜設計修正年月日及び修正ＩＤ＞                                              */
/*--------------------------------------------------------------------------*/
/* ＜ソース作成者名＞               Amos Xia                                  */
/* ＜ソース作成年月日＞             2020年06月20日　　　　　                    */
/* ＜ソース修正者名＞                                                         */
/* ＜ソース修正年月日及び修正ＩＤ＞                                            */
/*--------------------------------------------------------------------------*/

using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.IO;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;

namespace VMPrint
{
    class XpsPdfPrinter
    {
        /// <summary>
        /// Convert xps -> bitmap -> pdf.
        /// </summary>
        public static void ConvertXpsToBitmapToPdf(string xpsFilePath, string finalPdfPath)
        {
            MessageBox.Show("Before opening file: " + xpsFilePath, "XpsPDFPrinter",
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question);

            XpsDocument xps = new XpsDocument(xpsFilePath, FileAccess.Read);
            FixedDocumentSequence sequence = xps.GetFixedDocumentSequence();
            MessageBox.Show("Opened file: " + xpsFilePath, "XpsPDFPrinter", 
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question);

            PdfDocument outputPdf = new PdfDocument();
            for (int pageCount = 0; pageCount < sequence.DocumentPaginator.PageCount; ++pageCount)
            {
                DocumentPage page = sequence.DocumentPaginator.GetPage(pageCount);
                RenderTargetBitmap toBitmap = new RenderTargetBitmap((int)page.Size.Width, (int)page.Size.Height, 96, 96, System.Windows.Media.PixelFormats.Default);
                toBitmap.Render(page.Visual);

                // Get Bitmap.
                MemoryStream stream = new MemoryStream();
                BitmapEncoder encoder = new BmpBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(toBitmap));
                encoder.Save(stream);
                XImage xImage = XImage.FromStream(stream);
                xImage.Interpolate = false;

                // TODO: Add watermark.
                MessageBox.Show("Added page: " + (pageCount + 1), "XpsPDFPrinter", 
                                 MessageBoxButtons.YesNo,
                                 MessageBoxIcon.Question);

                // New PDF page.
                PdfPage newPdfPage = new PdfPage();
                outputPdf.Pages.Add(newPdfPage);
                XGraphics xgr = XGraphics.FromPdfPage(outputPdf.Pages[pageCount]);
                xgr.DrawImage(xImage, 0, 0);
            }

            // Save final PDF.
            MessageBox.Show("Will save file to: " + finalPdfPath, "XpsPDFPrinter",
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question);
            outputPdf.Save(finalPdfPath);
            outputPdf.Close();
            MessageBox.Show("Done saving file to: " + finalPdfPath, "XpsPDFPrinter", 
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question);
        }

        /// <summary>
        /// Convert xps -> pdf.
        /// </summary>
        public static void ConvertXpsToPdf(string xpsFilePath, string finalPdfPath)
        {
            // TODO
            //MemoryStream lMemoryStream = new MemoryStream();
            //Package package = Package.Open(lMemoryStream, FileMode.Create);
            //XpsDocument doc = new XpsDocument(package);
            //XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(doc);
            //writer.Write(dp);
            //doc.Close();
            //package.Close();
            //var pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(lMemoryStream);
        }

        /// <summary>
        /// Clean up temporary files.
        /// </summary>
        private static void Cleanup(IntPtr gsInstancePtr)
        {
            // TODO
        }
    }
}
