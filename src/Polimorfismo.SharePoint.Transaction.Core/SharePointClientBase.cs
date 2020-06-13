// Copyright 2020 Polimorfismo - José Mauro da Silva Sandy
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Utils;
using Polimorfismo.SharePoint.Transaction.Commands;
using Polimorfismo.SharePoint.Transaction.Resources;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Base class for the implementation of the instances that will establish 
    /// communication with SharePoint to performs the operations.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:14:44 PM</Date>
    public abstract class SharePointClientBase : IDisposable
    {
        #region Fields

        private readonly LinkedList<ISharePointCommand> _commandQueue = new LinkedList<ISharePointCommand>();

        private readonly SharePointBackgroundTasks _sharePointBackgroundTasks = new SharePointBackgroundTasks();

        #endregion

        #region Properties

        internal readonly SharePointListItemTracking Tracking = new SharePointListItemTracking();

        protected readonly IReadOnlyList<string> IgnorePropertiesInsertOrUpdate = new List<string>
        {
            SharePointConstants.FieldNameId,
            SharePointConstants.FieldNameFileRef,
            SharePointConstants.FieldNameCreated,
            SharePointConstants.FieldNameModified,
            SharePointConstants.FieldNameUIVersionString
        };

        #endregion

        #region Constructors / Finalizers

        protected SharePointClientBase()
        {
        }

        ~SharePointClientBase() => Dispose(false);

        #endregion

        #region Methods

        protected internal abstract Task DeleteItem<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new();

        protected internal abstract Task UpdateItem<TSharePointItem>(int id, IReadOnlyDictionary<string, object> fields)
            where TSharePointItem : ISharePointItem, new();

        protected internal abstract Task<int> InsertItem<TSharePointItem>(IReadOnlyDictionary<string, object> fields)
            where TSharePointItem : ISharePointItem, new();

        protected internal abstract Task<ICollection<TSharePointItem>> GetItems<TSharePointItem>(string viewXml)
            where TSharePointItem : ISharePointItem, new();

        protected TSharePointItem CreateSharePointItem<TSharePointItem>()
            where TSharePointItem : ISharePointItem, new() => SharePointItemFactory.Create<TSharePointItem>();

        public abstract Task<SharePointUser> GetUserByLogin(string login);

        public abstract Task<SharePointDocumentInfo> GetDocuments(string documentLibraryName, string fileRef);

        public async Task<TSharePointItem> GetItemById<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new()
        {
            var items = await GetItems<TSharePointItem>(string.Format(SharePointQueries.QueryItemById, id));
            return items.FirstOrDefault();
        }

        public void AddItem<TSharePointItem>(TSharePointItem sharePointItem)
            where TSharePointItem : ISharePointItem, new()
        {
            EnqueueCommand<SharePointInsertCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public void UpdateItem<TSharePointItem>(TSharePointItem sharePointItem)
            where TSharePointItem : ISharePointItem, new()
        {
            EnqueueCommand<SharePointUpdateCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public void DeleteItem<TSharePointItem>(TSharePointItem sharePointItem)
            where TSharePointItem : ISharePointItem, new()
        {
            EnqueueCommand<SharePointDeleteCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public async Task SaveChanges()
        {
            var undoStack = new Stack<ISharePointCommand>();

            try
            {
                _sharePointBackgroundTasks.Wait(30);

                if (!_sharePointBackgroundTasks.AllTasksCompletedSuccess())
                {
                    _sharePointBackgroundTasks.Cancel();
                    throw new SharePointException(SharePointErrorCode.PreparationCommandNotCompleted, SharePointMessages.ERR400);
                }

                while (_commandQueue.Count > 0)
                {
                    var command = _commandQueue.First();

                    await command.Execute();

                    undoStack.Push(command);

                    _commandQueue.RemoveFirst();
                }
            }
            catch (Exception ex)
            {
                await Rollback(undoStack);

                throw new SharePointException(SharePointErrorCode.SaveChanges, ex.Message, ex);
            }
            finally
            {
                Tracking.Clear();
                _commandQueue.Clear();
                _sharePointBackgroundTasks.Clear();
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
            where TSharePointItem : ISharePointItem, new()
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

            _sharePointBackgroundTasks.Action(() =>
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
                _sharePointBackgroundTasks?.Dispose();
            }
        }

        #endregion
    }
}
