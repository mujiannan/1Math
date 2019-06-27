using System;

namespace Correspondence
{
    public interface IReportor
    {
        Reportor Reportor { get; }
    }
    public class Reportor
    {
        public delegate void ChangeMessage(object sender, MessageEventArgs e);
        public delegate void ChangeProgress(object sender, ProgressEventArgs e);
        public event ChangeMessage MessageChange;
        public event ChangeProgress ProgressChange;
        private object _sender;
        public Reportor(object sender)
        {
            _sender = sender;
        }
        public void Report(string newMessage)
        {
            MessageChange?.Invoke(this, new MessageEventArgs(newMessage));
        }
        public void Report(double newProgress)
        {
            ProgressChange?.Invoke(this, new ProgressEventArgs(newProgress));
        }
        public class MessageEventArgs : EventArgs
        {
            private readonly string _newMessage;
            public MessageEventArgs(string newMessage)
            {
                _newMessage = newMessage;
            }
            public string NewMessage
            {
                get
                {
                    return _newMessage;
                }
            }
        }
        public class ProgressEventArgs : EventArgs
        {
            private readonly double _newProgress;
            public ProgressEventArgs(double newProgress)
            {
                _newProgress = newProgress;
            }
            public double NewProgress
            {
                get
                {
                    return _newProgress;
                }
            }
        }
    }
}
