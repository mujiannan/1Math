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
        public void ChangeMessage(object Sender, Reportor.MessageEventArgs messageEventArgs)
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
        public void ChangeProgress(object Sender, Reportor.ProgressEventArgs progressEventArgs)
        {
            int value = (int)(100 * progressEventArgs.NewProgress);
            if (this.IsDisposed)
            {
                return;
            }
            if (this.progressBar.InvokeRequired)
            {
                progressBar.BeginInvoke(new Action(() => { progressBar.Value = value; }));
            }
            else
            {
                progressBar.Value = value;
            }
        }
    }
}
