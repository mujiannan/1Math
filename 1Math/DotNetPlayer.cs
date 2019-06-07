using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Threading;
namespace _1Math
{
    public class DotNetPlayer : IDisposable
    {
        private bool disposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        ~DotNetPlayer()
        {
            Dispose(false);
        }
        private void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //呵呵，暂时还没有啥托管资源需要释放的
                }
                mediaPlayer.Close();
                mediaPlayer = null;
                disposed = true;
            }
        }
        MediaPlayer mediaPlayer;
        const double timeOut = 5;
        public DotNetPlayer()
        {
            mediaPlayer = new MediaPlayer();
        }
        public double GetDuration(Uri uri)
        {
            mediaPlayer.Open(uri);
            double duration = 0;
            DateTime start = DateTime.Now;
            TimeSpan timeSpan;
            do
            {
                Thread.Sleep(50);
                if (mediaPlayer.NaturalDuration.HasTimeSpan)
                {
                    duration = mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                    mediaPlayer.Stop();
                }
                timeSpan = DateTime.Now - start;
            } while (duration == 0 && timeSpan.TotalSeconds < timeOut);
            return (duration);
        }
    }
}
