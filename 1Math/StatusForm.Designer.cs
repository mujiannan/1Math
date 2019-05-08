using System;

namespace _1Math
{
    partial class StatusForm
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
            this.MessageLabel = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // MessageLabel
            // 
            this.MessageLabel.AutoSize = true;
            this.MessageLabel.Location = new System.Drawing.Point(19, 14);
            this.MessageLabel.Name = "MessageLabel";
            this.MessageLabel.Size = new System.Drawing.Size(0, 12);
            this.MessageLabel.TabIndex = 0;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(35, 44);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(513, 23);
            this.progressBar1.TabIndex = 1;
            this.progressBar1.Click += new System.EventHandler(this.ProgressBar1_Click);
            // 
            // StatusForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 91);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.MessageLabel);
            this.Name = "StatusForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "处理状态";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.StatusForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        public void MessageLabel_TextChange(object Sender,MessageEventArgs messageEventArgs)
        {
            if (this.IsDisposed)
            {
                return;
            }
            if (this.MessageLabel.InvokeRequired)
            {
                MessageLabel.BeginInvoke(new Action(() => { MessageLabel.Text = messageEventArgs.NewMessage; }));
            }
            else
            {
                MessageLabel.Text = messageEventArgs.NewMessage;
            }
        }
        public void ProgressBar_ValueChange(object Sender,ProgressEventArgs progressEventArgs)
        {
            int value = (int)(100 * progressEventArgs.NewProgress);
            if (this.IsDisposed)
            {
                return;
            }
            if (this.progressBar1.InvokeRequired)
            {
                progressBar1.BeginInvoke(new Action(() => { progressBar1.Value = value; }));
            }
            else
            {
                progressBar1.Value =value;
            }
        }
        public System.Windows.Forms.Label MessageLabel;
        public System.Windows.Forms.ProgressBar progressBar1;
    }
}