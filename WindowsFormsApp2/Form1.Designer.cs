﻿
namespace WindowsFormsApp2
{
    partial class Form1
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
            this.btnGenerate = new System.Windows.Forms.Button();
            this.lState = new System.Windows.Forms.Label();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.InputFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(384, 12);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 0;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // lState
            // 
            this.lState.AutoSize = true;
            this.lState.Location = new System.Drawing.Point(9, 37);
            this.lState.Name = "lState";
            this.lState.Size = new System.Drawing.Size(16, 13);
            this.lState.TabIndex = 1;
            this.lState.Text = "...";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(303, 12);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 3;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // InputFile
            // 
            this.InputFile.Location = new System.Drawing.Point(12, 14);
            this.InputFile.Name = "InputFile";
            this.InputFile.Size = new System.Drawing.Size(285, 20);
            this.InputFile.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(296, 26);
            this.label1.TabIndex = 5;
            this.label1.Text = "1. Click \"Browse\" to select input file (must be sorted by name)\r\n2. Click \"Genera" +
    "te\" to generate output file";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(468, 92);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.InputFile);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.lState);
            this.Controls.Add(this.btnGenerate);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Data Processing";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Label lState;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox InputFile;
        private System.Windows.Forms.Label label1;
    }
}

