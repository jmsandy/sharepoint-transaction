using System;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace Polimorfismo.SharePoint.Transaction.Core
{
    /// <summary>
    /// Controls the background tasks that are performed for the preparation of commands.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-25 11:27:26 PM</Date>
    internal class SharePointBackgroundTasks : IDisposable
    {
        #region Properties

        internal CancellationToken Token => CancellationTokenSource.Token;

        private ConcurrentBag<Task> Tasks { get; } = new ConcurrentBag<Task>();

        private CancellationTokenSource CancellationTokenSource { get; } = new CancellationTokenSource();

        #endregion

        #region Constructors / Finalizers

        public SharePointBackgroundTasks()
        {
        }

        ~SharePointBackgroundTasks() => Dispose(false);

        #endregion

        #region Methods

        public void Action(Action action)
        {
            Tasks.Add(Task.Run(action, Token));
        }

        public void Wait(int timeout)
        {
            _ = Task.WhenAny(Task.WhenAll(Tasks), Task.Delay(TimeSpan.FromSeconds(timeout))).Result;
        }

        public void Cancel()
        {
            CancellationTokenSource.Cancel();
        }

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (disposing)
            {
                Cancel();
            }
        }

        #endregion
    }
}
