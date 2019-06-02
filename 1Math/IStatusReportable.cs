using System;

namespace _1Math
{
    public delegate void ChangeMessage(object Sender, MessageEventArgs messageEventArgs);
    public delegate void ChangeProgress(object Sender, ProgressEventArgs progressEventArgs);
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
    public interface IHasStatusReporter
    {
        event ChangeMessage MessageChange;
        event ChangeProgress ProgressChange;
    }
}
