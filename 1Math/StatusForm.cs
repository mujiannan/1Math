using System;
using System.Windows.Forms;

namespace _1Math
{
    public partial class StatusForm : Form
    {

        public StatusForm()
        {
            InitializeComponent();
        }
        private void StatusForm_Load(object sender, EventArgs e)
        {

        }

        private void ProgressBar1_Click(object sender, EventArgs e)
        {

        }
        public void MessageLabel_TextChange(object Sender, Reportor.MessageEventArgs messageEventArgs)
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
        public void ProgressBar_ValueChange(object Sender, Reportor.ProgressEventArgs progressEventArgs)
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
                progressBar1.Value = value;
            }
        }
    }
}
