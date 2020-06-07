using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Core.Commands;

namespace Polimorfismo.SharePoint.Transaction.Core
{
    /// <summary>
    /// Base class for the implementation of the instances that will establish 
    /// communication with SharePoint to performs the operations.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:14:44 PM</Date>
    public abstract class SharePointClientBase : IDisposable
    {
        #region Properties

        internal readonly SharePointListItemTracking Tracking = new SharePointListItemTracking();

        private readonly LinkedList<ISharePointCommand> _commandQueue = new LinkedList<ISharePointCommand>();

        private SharePointBackgroundTasks SharePointBackgroundTasks = new SharePointBackgroundTasks();

        #endregion

        #region Constructors / Finalizers

        /// <summary>
        /// 
        /// </summary>
        protected SharePointClientBase()
        {
        }

        ~SharePointClientBase() => Dispose(false);

        #endregion

        #region Methods

        protected internal abstract Task DeleteItem<TSharePointItem>(int id) 
            where TSharePointItem : ISharePointItem;

        protected internal abstract Task UpdateItem<TSharePointItem>(int id, IReadOnlyDictionary<string, object> fields) 
            where TSharePointItem : ISharePointItem;

        protected internal abstract Task<int> InsertItem<TSharePointItem>(IReadOnlyDictionary<string, object> fields) 
            where TSharePointItem : ISharePointItem;

        protected internal abstract Task<ICollection<TSharePointItem>> GetItems<TSharePointItem>(string viewXml) 
            where TSharePointItem : ISharePointItem;

        public async Task<TSharePointItem> GetItemById<TSharePointItem>(int id) 
            where TSharePointItem : ISharePointItem
        {
            var items = await GetItems<TSharePointItem>($"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>{id}</Value></Eq></Where></Query></View>");
            return items.FirstOrDefault();
        }

        public void AddItem<TSharePointItem>(TSharePointItem sharePointItem) 
            where TSharePointItem : ISharePointItem
        {
            EnqueueCommand<SharePointInsertCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public void UpdateItem<TSharePointItem>(TSharePointItem sharePointItem) 
            where TSharePointItem : ISharePointItem
        {
            EnqueueCommand<SharePointUpdateCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public void DeleteItem<TSharePointItem>(TSharePointItem sharePointItem) 
            where TSharePointItem : ISharePointItem
        {
            EnqueueCommand<SharePointDeleteCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public async Task SaveChanges()
        {
            var undoStack = new Stack<ISharePointCommand>();

            SharePointBackgroundTasks.Wait(60);
            SharePointBackgroundTasks.Cancel();

#warning Check running tasks
            try
            {
                while (_commandQueue.Count > 0)
                {
                    var command = _commandQueue.First();

                    await command.Execute();

                    undoStack.Push(command);

                    _commandQueue.RemoveFirst();
                }
            }
            catch
            {
                await Rollback(undoStack);
            }
        }

        private async Task Rollback(Stack<ISharePointCommand> undoStack)
        {
            while (undoStack.Count > 0)
            {
                await undoStack.Pop().Undo();
            }
        }

        private void EnqueueCommand<TSharePointCommand, TSharePointItem>(ISharePointItem sharePointItem)
            where TSharePointItem : ISharePointItem
            where TSharePointCommand : SharePointCommand<TSharePointItem>
        {
            ISharePointCommand referenceCommand = null;
            foreach (var itemTracking in Tracking.Items)
            {
                if (itemTracking.Item.GetReferences()
                    .Any(r => ReferenceEquals(r.Value, sharePointItem)))
                {
                    referenceCommand = _commandQueue.First(c => ReferenceEquals(c.SharePointItemTracking.Item, itemTracking.Item));
                    break;
                }
            }

            var trackingItem = new SharePointItemTracking(sharePointItem);
            Tracking.Add(trackingItem);

            var command = (TSharePointCommand)Activator.CreateInstance(typeof(TSharePointCommand), this, trackingItem);

            if (referenceCommand != null)
            {
                _commandQueue.AddBefore(_commandQueue.Find(referenceCommand), command);
            }
            else
            {
                _commandQueue.AddLast(command);
            }

            SharePointBackgroundTasks.Action(() =>
            {
                command.Prepare().Wait();
            });
        }

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                Tracking?.Dispose();
                _commandQueue?.Clear();
                SharePointBackgroundTasks?.Dispose();
            }
        }

        #endregion
    }
}
