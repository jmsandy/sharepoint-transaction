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
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Polimorfismo.SharePoint.Transaction.Utils;
using Polimorfismo.SharePoint.Transaction.Commands;
using Polimorfismo.SharePoint.Transaction.Resources;

[assembly: InternalsVisibleTo("Polimorfismo.SharePointOnline.Transaction.CSOM.Tests")]

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

        public SharePointUser GetUserByLogin(string login)
        {
            return GetUserByLoginAsync(login).GetAwaiter().GetResult();
        }

        public abstract Task<SharePointUser> GetUserByLoginAsync(string login);

        protected internal abstract Task<int> AddItemAsync<TSharePointItem>(IReadOnlyDictionary<string, object> fields)
            where TSharePointItem : ISharePointItem, new();

        protected internal abstract Task UpdateItemAsync<TSharePointItem>(int id, IReadOnlyDictionary<string, object> fields)
            where TSharePointItem : ISharePointItem, new();

        protected internal abstract Task DeleteItemAsync<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new();

        protected TSharePointMetadata CreateSharePointItem<TSharePointMetadata>()
             where TSharePointMetadata : ISharePointMetadata, new() => SharePointItemFactory.Create<TSharePointMetadata>();

        protected internal abstract Task<ICollection<TSharePointMetadata>> GetItemsAsync<TSharePointMetadata>(string viewXml)
            where TSharePointMetadata : ISharePointMetadata, new();

        protected internal abstract Task<(int Id, List<string> CreatedFolders)> AddFileAsync<TSharePointFile>(IReadOnlyDictionary<string, object> fields, 
            string fileName, string folderName, Stream content, bool isUpdateFile) where TSharePointFile : ISharePointFile, new();

        protected internal abstract Task DeleteFileAsync<TSharePointFile>(int id)
            where TSharePointFile : ISharePointFile, new();

        protected internal abstract Task RemoveFoldersAsync<TSharePointFile>(List<string> folders)
            where TSharePointFile : ISharePointFile, new();

        public SharePointDocumentInfo GetFiles(string documentLibraryName, string fileRef)
        {
            return GetFilesAsync(documentLibraryName, fileRef).GetAwaiter().GetResult();
        }

        public abstract Task<SharePointDocumentInfo> GetFilesAsync(string documentLibraryName, string fileRef);

        public TSharePointFile GetFileById<TSharePointFile>(int id)
            where TSharePointFile : ISharePointFile, new()
        {
            return GetFileByIdAsync<TSharePointFile>(id).GetAwaiter().GetResult();
        }

        public async Task<TSharePointFile> GetFileByIdAsync<TSharePointFile>(int id)
            where TSharePointFile : ISharePointFile, new()
        {
            var files = await GetItemsAsync<TSharePointFile>(string.Format(SharePointQueries.QueryItemById, id));
            return files.FirstOrDefault();
        }

        public TSharePointItem GetItemById<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new()
        {
            return GetItemByIdAsync<TSharePointItem>(id).GetAwaiter().GetResult();
        }

        public async Task<TSharePointItem> GetItemByIdAsync<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new()
        {
            var items = await GetItemsAsync<TSharePointItem>(string.Format(SharePointQueries.QueryItemById, id));
            return items.FirstOrDefault();
        }

        public void AddItem<TSharePointItem>(TSharePointItem sharePointItem)
            where TSharePointItem : ISharePointItem, new()
        {
            EnqueueCommand<SharePointAddItemCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public void UpdateItem<TSharePointItem>(TSharePointItem sharePointItem)
            where TSharePointItem : ISharePointItem, new()
        {
            EnqueueCommand<SharePointUpdateItemCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public void DeleteItem<TSharePointItem>(TSharePointItem sharePointItem)
            where TSharePointItem : ISharePointItem, new()
        {
            EnqueueCommand<SharePointDeleteItemCommand<TSharePointItem>, TSharePointItem>(sharePointItem);
        }

        public void AddFile<TSharePointFile>(TSharePointFile sharePointFile)
            where TSharePointFile : ISharePointFile, new()
        {
            EnqueueCommand<SharePointAddFileCommand<TSharePointFile>, TSharePointFile>(sharePointFile);
        }

        public void UpdateFile<TSharePointFile>(TSharePointFile sharePointFile)
            where TSharePointFile : ISharePointFile, new()
        {
            EnqueueCommand<SharePointUpdateFileCommand<TSharePointFile>, TSharePointFile>(sharePointFile);
        }

        public void DeleteFile<TSharePointFile>(TSharePointFile sharePointFile)
            where TSharePointFile : ISharePointFile, new()
        {
            EnqueueCommand<SharePointDeleteFileCommand<TSharePointFile>, TSharePointFile>(sharePointFile);
        }

        public void SaveChanges()
        {
            SaveChangesAsync().GetAwaiter().GetResult();
        }

        public async Task SaveChangesAsync()
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

                    await command.ExecuteAsync();

                    undoStack.Push(command);

                    _commandQueue.RemoveFirst();
                }
            }
            catch (Exception ex)
            {
                await RollbackAsync(undoStack);

                throw new SharePointException(SharePointErrorCode.SaveChanges, ex.Message, ex);
            }
            finally
            {
                Tracking.Clear();
                _commandQueue.Clear();
                _sharePointBackgroundTasks.Clear();
            }
        }

        private async Task RollbackAsync(Stack<ISharePointCommand> undoStack)
        {
            while (undoStack.Count > 0)
            {
                await undoStack.Pop().UndoAsync();
            }
        }

        private void EnqueueCommand<TSharePointCommand, TSharePointMetadata>(ISharePointMetadata sharePointItem)
            where TSharePointMetadata : ISharePointMetadata, new()
            where TSharePointCommand : SharePointCommand<TSharePointMetadata>
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
                command.PrepareAsync().GetAwaiter().GetResult();
                _sharePointBackgroundTasks.Token.ThrowIfCancellationRequested();
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
