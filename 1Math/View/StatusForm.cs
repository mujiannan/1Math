using System;
using System.Windows.Forms;
using Correspondence;
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
        public void SetStyle(ProgressBarStyle style)
        {
            progressBar?.Invoke(new Action(() => progressBar.Style = style));
        }
        public void ChangeMessage(object Sender, Reportor.MessageEventArgs messageEventArgs)
        {
            MessageLabel?.Invoke(new Action(() => MessageLabel.Text = messageEventArgs.NewMessage));
        }
        public void ChangeProgress(object Sender, Reportor.ProgressEventArgs progressEventArgs)
        {
            progressBar?.Invoke(new Action(() =>progressBar.Value =(int)(100 * progressEventArgs.NewProgress)));
        }
    }
}
