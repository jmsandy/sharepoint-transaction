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
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Polimorfismo.SharePoint.Transaction.Utils;
using Polimorfismo.SharePoint.Transaction.Commands;
using Polimorfismo.SharePoint.Transaction.Resources;

[assembly: InternalsVisibleTo("Polimorfismo.SharePoint.Transaction.Core.Tests")]
[assembly: InternalsVisibleTo("Polimorfismo.SharePointOnline.Transaction.CSOM")]
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

        protected bool Disposed { get; private set; }

        internal readonly SharePointListItemTracking Tracking = new SharePointListItemTracking();

        protected readonly IReadOnlyList<string> IgnorePropertiesInsertOrUpdate = new List<string>
        {
            SharePointConstants.FieldNameId,
            SharePointConstants.FieldNameBSN,
            SharePointConstants.FieldNameGUID,
            SharePointConstants.FieldNameDirty,
            SharePointConstants.FieldNameOrder,
            SharePointConstants.FieldNameLevel,
            SharePointConstants.FieldNameProgId,
            SharePointConstants.FieldNameFileRef,
            SharePointConstants.FieldNameCreated,
            SharePointConstants.FieldNameScopeId,
            SharePointConstants.FieldNameModified,
            SharePointConstants.FieldNameFileType,
            SharePointConstants.FieldNameMetaInfo,
            SharePointConstants.FieldNameFileSize,
            SharePointConstants.FieldNameXdProgID,
            SharePointConstants.FieldNameParsable,
            SharePointConstants.FieldNameStubFile,
            SharePointConstants.FieldNameUniqueId,
            SharePointConstants.FieldNameSourceUrl,
            SharePointConstants.FieldNameLikeCount,
            SharePointConstants.FieldNameIpLabelId,
            SharePointConstants.FieldNameNoExecute,
            SharePointConstants.FieldNameFSObjType,
            SharePointConstants.FieldNameAppAuthor,
            SharePointConstants.FieldNameAppEditor,
            SharePointConstants.FieldNameVirusInfo,
            SharePointConstants.FieldNameUIVersion,
            SharePointConstants.FieldNameStreamHash,
            SharePointConstants.FieldNameRestricted,
            SharePointConstants.FieldNameInstanceID,
            SharePointConstants.FieldNameCopySource,
            SharePointConstants.FieldNameFileDirRef,
            SharePointConstants.FieldNameTemplateUrl,
            SharePointConstants.FieldNameXdSignature,
            SharePointConstants.FieldNameAttachments,
            SharePointConstants.FieldNameSMTotalSize,
            SharePointConstants.FieldNameVirusStatus,
            SharePointConstants.FieldNameCreatedDate,
            SharePointConstants.FieldNameFileLeafRef,
            SharePointConstants.FieldNameShortcutUrl,
            SharePointConstants.FieldNameDisplayName,
            SharePointConstants.FieldNameLastModified,
            SharePointConstants.FieldNameOriginatorId,
            SharePointConstants.FieldNameSyncClientId,
            SharePointConstants.FieldNameSortBehavior,
            SharePointConstants.FieldNameAccessPolicy,
            SharePointConstants.FieldNameHTMLFileType,
            SharePointConstants.FieldNameCommentFlags,
            SharePointConstants.FieldNameCommentCount,
            SharePointConstants.FieldNameVirusStatus2,
            SharePointConstants.FieldNameCheckoutUser,
            SharePointConstants.FieldNameRmsTemplateId,
            SharePointConstants.FieldNameShortcutWebId,
            SharePointConstants.FieldNameVirusVendorID,
            SharePointConstants.FieldNameContentTypeId,
            SharePointConstants.FieldNameComplianceTag,
            SharePointConstants.FieldNameA2ODMountCount,
            SharePointConstants.FieldNameParentLeafName,
            SharePointConstants.FieldNameParentUniqueId,
            SharePointConstants.FieldNameShortcutSiteId,
            SharePointConstants.FieldNameCheckinComment,
            SharePointConstants.FieldNameContentVersion,
            SharePointConstants.FieldNameItemChildCount,
            SharePointConstants.FieldNameUIVersionString,
            SharePointConstants.FieldNameWorkflowVersion,
            SharePointConstants.FieldNameComplianceFlags,
            SharePointConstants.FieldNameCheckedOutTitle,
            SharePointConstants.FieldNameMediaServiceOCR,
            SharePointConstants.FieldNameSharedFileIndex,
            SharePointConstants.FieldNameCheckedOutUserId,
            SharePointConstants.FieldNameShortcutUniqueId,
            SharePointConstants.FieldNameSMTotalFileCount,
            SharePointConstants.FieldNameFolderChildCount,
            SharePointConstants.FieldNameOwshiddenversion,
            SharePointConstants.FieldNameIsCurrentVersion,
            SharePointConstants.FieldNameModerationStatus,
            SharePointConstants.FieldNameListSchemaVersion,
            SharePointConstants.FieldNameComplianceAssetId,
            SharePointConstants.FieldNameSMLastModifiedDate,
            SharePointConstants.FieldNameWorkflowInstanceID,
            SharePointConstants.FieldNameModerationComments,
            SharePointConstants.FieldNameHasCopyDestinations,
            SharePointConstants.FieldNameComplianceTagUserId,
            SharePointConstants.FieldNameIsCheckedoutToLocal,
            SharePointConstants.FieldNameParentVersionString,
            SharePointConstants.FieldNameHasEncryptedContent,
            SharePointConstants.FieldNameDocConcurrencyNumber,
            SharePointConstants.FieldNameMediaServiceAutoTags,
            SharePointConstants.FieldNameMediaServiceMetadata,
            SharePointConstants.FieldNameSMTotalFileStreamSize,
            SharePointConstants.FieldNameIpLabelAssignmentMethod,
            SharePointConstants.FieldNameComplianceTagWrittenTime,
            SharePointConstants.FieldNameMediaServiceFastMetadata,
            SharePointConstants.FieldNameMediaServiceEventHashCode,
            SharePointConstants.FieldNameMediaServiceGenerationTime
        };

        #endregion

        #region Constructors / Finalizers

        protected SharePointClientBase()
        {
        }

        ~SharePointClientBase() => Dispose(false);

        #endregion

        #region Methods
        public abstract SharePointUser GetUserByLogin(string login);

        public abstract Task<SharePointUser> GetUserByLoginAsync(string login);

        protected internal abstract Task<int> AddItemAsync<TSharePointItem>(IReadOnlyDictionary<string, object> fields)
            where TSharePointItem : ISharePointItem, new();

        protected internal abstract Task UpdateItemAsync<TSharePointItem>(int id, IReadOnlyDictionary<string, object> fields)
            where TSharePointItem : ISharePointItem, new();

        protected internal abstract Task DeleteItemAsync<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new();

        protected TSharePointMetadata CreateSharePointItem<TSharePointMetadata>()
             where TSharePointMetadata : ISharePointMetadata, new() => SharePointItemFactory.Create<TSharePointMetadata>();

        protected internal abstract Task<KeyValuePair<ISharePointMetadata, Dictionary<string, object>>> GetAllFieldsByIdAsync<TSharePointMetadata>(int id)
            where TSharePointMetadata : ISharePointMetadata, new();

        protected internal abstract Task<ICollection<TSharePointMetadata>> GetItemsAsync<TSharePointMetadata>(string viewXml)
            where TSharePointMetadata : ISharePointMetadata, new();

        protected internal abstract Task<(int Id, List<string> CreatedFolders)> AddFileAsync<TSharePointFile>(IReadOnlyDictionary<string, object> fields, 
            string fileName, string folderName, Stream content, bool isUpdateFile) where TSharePointFile : ISharePointFile, new();

        protected internal abstract Task DeleteFileAsync<TSharePointFile>(int id)
            where TSharePointFile : ISharePointFile, new();

        protected internal abstract Task RemoveFoldersAsync<TSharePointFile>(List<string> folders)
            where TSharePointFile : ISharePointFile, new();

        public SharePointDocumentInfo GetDocumentsInfo(string documentLibraryName, string fileRef)
        {
            return Task<Task<SharePointDocumentInfo>>.Factory.StartNew(() => GetDocumentsInfoAsync(documentLibraryName, fileRef),
                CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).Unwrap().GetAwaiter().GetResult();
        }

        public abstract Task<SharePointDocumentInfo> GetDocumentsInfoAsync(string documentLibraryName, string fileRef);

        public TSharePointFile GetFileById<TSharePointFile>(int id)
            where TSharePointFile : ISharePointFile, new()
        {
            return Task<Task<TSharePointFile>>.Factory.StartNew(() => GetFileByIdAsync<TSharePointFile>(id), 
                CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).Unwrap().GetAwaiter().GetResult();
        }

        public async Task<TSharePointFile> GetFileByIdAsync<TSharePointFile>(int id)
            where TSharePointFile : ISharePointFile, new()
        {
            return (await GetItemsAsync<TSharePointFile>(string.Format(SharePointQueries.QueryItemById, id))).FirstOrDefault();
        }

        public TSharePointItem GetItemById<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new()
        {
            return Task<Task<TSharePointItem>>.Factory.StartNew(() => GetItemByIdAsync<TSharePointItem>(id),
                CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).Unwrap().GetAwaiter().GetResult();
        }

        public async Task<TSharePointItem> GetItemByIdAsync<TSharePointItem>(int id)
            where TSharePointItem : ISharePointItem, new()
        {
            return (await GetItemsAsync<TSharePointItem>(string.Format(SharePointQueries.QueryItemById, id))).FirstOrDefault();
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
            Task.Factory.StartNew(() => SaveChangesAsync(),
                CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).Unwrap().GetAwaiter().GetResult();
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

                    if (command.StackFirst) undoStack.Push(command);

                    await command.ExecuteAsync();

                    if (!command.StackFirst) undoStack.Push(command);

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
                Task.Factory.StartNew(() => command.PrepareAsync(),
                    CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).Unwrap().GetAwaiter().GetResult();
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

        protected virtual void Dispose(bool disposing)
        {
            if (Disposed) return;

            if (disposing)
            {
                Tracking?.Dispose();
                _commandQueue?.Clear();
                _sharePointBackgroundTasks?.Dispose();
            }

            Disposed = true;
        }

        #endregion
    }
}
