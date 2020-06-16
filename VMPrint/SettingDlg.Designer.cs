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
namespace VMPrint
{
    partial class SettingDlg
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtBranchNo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.txtMachineNo = new System.Windows.Forms.TextBox();
            this.txtRealPrinterName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtOutputDir = new System.Windows.Forms.TextBox();
            this.txtStyle = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtSerialNo = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtBranchNo
            // 
            this.txtBranchNo.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.txtBranchNo.Location = new System.Drawing.Point(231, 63);
            this.txtBranchNo.MaxLength = 4;
            this.txtBranchNo.Name = "txtBranchNo";
            this.txtBranchNo.Size = new System.Drawing.Size(100, 25);
            this.txtBranchNo.TabIndex = 1;
            this.txtBranchNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBranchNo_KeyPress);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(156, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 4;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(498, 456);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(105, 29);
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // txtMachineNo
            // 
            this.txtMachineNo.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.txtMachineNo.Location = new System.Drawing.Point(231, 95);
            this.txtMachineNo.MaxLength = 3;
            this.txtMachineNo.Name = "txtMachineNo";
            this.txtMachineNo.Size = new System.Drawing.Size(100, 25);
            this.txtMachineNo.TabIndex = 2;
            this.txtMachineNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMachineNo_KeyPress);
            // 
            // txtRealPrinterName
            // 
            this.txtRealPrinterName.Location = new System.Drawing.Point(231, 32);
            this.txtRealPrinterName.Name = "txtRealPrinterName";
            this.txtRealPrinterName.Size = new System.Drawing.Size(420, 25);
            this.txtRealPrinterName.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(27, 35);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(181, 15);
            this.label4.TabIndex = 5;
            this.label4.Text = "出力ディレクトリ：";
            this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtOutputDir
            // 
            this.txtOutputDir.Location = new System.Drawing.Point(231, 32);
            this.txtOutputDir.Name = "txtOutputDir";
            this.txtOutputDir.Size = new System.Drawing.Size(420, 25);
            this.txtOutputDir.TabIndex = 0;
            // 
            // txtStyle
            // 
            this.txtStyle.Location = new System.Drawing.Point(231, 162);
            this.txtStyle.Multiline = true;
            this.txtStyle.Name = "txtStyle";
            this.txtStyle.Size = new System.Drawing.Size(420, 146);
            this.txtStyle.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.txtSerialNo);
            this.groupBox1.Controls.Add(this.txtStyle);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtOutputDir);
            this.groupBox1.Controls.Add(this.txtBranchNo);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtMachineNo);
            this.groupBox1.Location = new System.Drawing.Point(22, 22);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(702, 328);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "PDF";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(27, 165);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(181, 15);
            this.label3.TabIndex = 7;
            this.label3.Text = "検索文字列：";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label8.Location = new System.Drawing.Point(156, 98);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(52, 15);
            this.label8.TabIndex = 9;
            this.label8.Text = "機番：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label7.Location = new System.Drawing.Point(156, 66);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(52, 15);
            this.label7.TabIndex = 8;
            this.label7.Text = "店番：";
            this.label7.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtSerialNo
            // 
            this.txtSerialNo.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.txtSerialNo.Location = new System.Drawing.Point(231, 126);
            this.txtSerialNo.MaxLength = 3;
            this.txtSerialNo.Name = "txtSerialNo";
            this.txtSerialNo.ReadOnly = true;
            this.txtSerialNo.Size = new System.Drawing.Size(100, 25);
            this.txtSerialNo.TabIndex = 3;
            this.txtSerialNo.TabStop = false;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label6.Location = new System.Drawing.Point(156, 129);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 15);
            this.label6.TabIndex = 1;
            this.label6.Text = "連番：";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.txtRealPrinterName);
            this.groupBox2.Location = new System.Drawing.Point(22, 369);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(702, 71);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "通常プリンター";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(27, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(181, 15);
            this.label2.TabIndex = 6;
            this.label2.Text = "プリンター名：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(609, 456);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(105, 29);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "閉じる";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // SettingDlg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(751, 497);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnSave);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingDlg";
            this.Text = "パラメータ設定";
            this.Load += new System.EventHandler(this.SettingDlg_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox txtBranchNo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtMachineNo;
        private System.Windows.Forms.TextBox txtRealPrinterName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtOutputDir;
        private System.Windows.Forms.TextBox txtStyle;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtSerialNo;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}