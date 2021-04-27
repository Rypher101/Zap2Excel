namespace Zap2Excel
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
            this.txtHTML = new System.Windows.Forms.TextBox();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btnHTML = new System.Windows.Forms.Button();
            this.btnOutput = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.txtLog = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // txtHTML
            // 
            this.txtHTML.Location = new System.Drawing.Point(12, 49);
            this.txtHTML.Name = "txtHTML";
            this.txtHTML.Size = new System.Drawing.Size(237, 20);
            this.txtHTML.TabIndex = 0;
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(12, 109);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(237, 20);
            this.txtOutput.TabIndex = 1;
            // 
            // btnHTML
            // 
            this.btnHTML.Location = new System.Drawing.Point(255, 47);
            this.btnHTML.Name = "btnHTML";
            this.btnHTML.Size = new System.Drawing.Size(120, 23);
            this.btnHTML.TabIndex = 3;
            this.btnHTML.Text = "Browse  HTML File";
            this.btnHTML.UseVisualStyleBackColor = true;
            this.btnHTML.Click += new System.EventHandler(this.btnHTML_Click);
            // 
            // btnOutput
            // 
            this.btnOutput.Location = new System.Drawing.Point(255, 106);
            this.btnOutput.Name = "btnOutput";
            this.btnOutput.Size = new System.Drawing.Size(120, 23);
            this.btnOutput.TabIndex = 4;
            this.btnOutput.Text = "Browse Output Location";
            this.btnOutput.UseVisualStyleBackColor = true;
            this.btnOutput.Click += new System.EventHandler(this.btnOutput_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(645, 49);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(143, 80);
            this.btnStart.TabIndex = 5;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(12, 162);
            this.txtLog.Name = "txtLog";
            this.txtLog.Size = new System.Drawing.Size(776, 263);
            this.txtLog.TabIndex = 7;
            this.txtLog.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(799, 437);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnOutput);
            this.Controls.Add(this.btnHTML);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.txtHTML);
            this.Name = "Form1";
            this.Text = "Zap 2 Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtHTML;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btnHTML;
        private System.Windows.Forms.Button btnOutput;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.RichTextBox txtLog;
    }
}

