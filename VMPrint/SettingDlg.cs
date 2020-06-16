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
/*〈モジュール名〉                  SettingDlg                              */
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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VMPrint
{
    public partial class SettingDlg : Form
    {
        public SettingDlg()
        {
            InitializeComponent();
        }

        private void SettingDlg_Load(object sender, EventArgs e)
        {
            txtOutputDir.Text = Properties.Settings.Default.OutputDir;
            txtBranchNo.Text = Properties.Settings.Default.BranchNo;
            txtMachineNo.Text = Properties.Settings.Default.MachineNo;
            txtSerialNo.Text = Convert.ToString(Properties.Settings.Default.SerialNo);
            txtRealPrinterName.Text = Properties.Settings.Default.RealPrinterName;
            txtStyle.Text = Properties.Settings.Default.Style;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtOutputDir.Text.Length == 0)
            {
                MessageBox.Show(Properties.Resources.ERROR_PDF_DIRECTORY_NOTINPUT, Properties.Resources.ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (Program.CreateDirectory(Environment.ExpandEnvironmentVariables(txtOutputDir.Text)) == false)
            {
                MessageBox.Show(Properties.Resources.ERROR_PDF_DIRECTORY_INVALID, Properties.Resources.ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (txtBranchNo.Text.Length != 4)
            {
                MessageBox.Show(Properties.Resources.ERROR_BRANCH_NO, Properties.Resources.ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (txtMachineNo.Text.Length != 3)
            {
                MessageBox.Show(Properties.Resources.ERROR_MACHINE_NO, Properties.Resources.ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (txtStyle.Text.Length == 0)
            {
                MessageBox.Show(Properties.Resources.ERROR_STYLE_NOTINPUT, Properties.Resources.ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (txtRealPrinterName.Text.Length == 0)
            {
                MessageBox.Show(Properties.Resources.ERROR_PRINT_NAME_NOTINPUT, Properties.Resources.ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Properties.Settings.Default.OutputDir = txtOutputDir.Text;
            Properties.Settings.Default.BranchNo = txtBranchNo.Text;
            Properties.Settings.Default.MachineNo = txtMachineNo.Text;
            Properties.Settings.Default.RealPrinterName = txtRealPrinterName.Text;
            Properties.Settings.Default.Style = txtStyle.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtBranchNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && !Char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtMachineNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= 'a' && e.KeyChar <= 'z') ||
                (e.KeyChar >= 'A' && e.KeyChar <= 'Z') ||
                (e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar == 8))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }
    }
}
