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
using System.Net;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Runtime.CompilerServices;
using Polimorfismo.SharePoint.Transaction.Utils;
using Polimorfismo.SharePoint.Transaction.Resources;

[assembly: InternalsVisibleTo("Polimorfismo.SharePointOnline.Transaction.CSOM.Tests")]

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Client context to perform operations in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-25 09:26:12 PM</Date>
    public class SharePointClient : SharePointClientBase
    {
        #region Attributes

        private ClientContext _clientContext;

        private readonly string _webFullUrl;

        private readonly object _lock = new object();

        private readonly ICredentials _networkCredential;

        protected readonly IList<SharePointUser> _users = new List<SharePointUser>();

        #endregion

        #region Properties

        private ClientContext ClientContext
        {
            get
            {
                if (_clientContext == null)
                {
                    _clientContext = new ClientContext(_webFullUrl)
                    {
                        Credentials = _networkCredential
                    };
                }

                return _clientContext;
            }
        }

        #endregion

        #region Constructors/Finalizers

        public SharePointClient(string webFullUrl, ICredentials credential)
            : base()
        {
            _webFullUrl = webFullUrl;
            _networkCredential = credential;
        }

        ~SharePointClient() => Dispose(false);

        #endregion

        #region SharePointClientBase - Members

        public override async Task<SharePointUser> GetUserByLoginAsync(string login)
        {
            return await Task.Run(() => GetUserByLogin(login));
        }

        public override SharePointUser GetUserByLogin(string login)
        {
            if (string.IsNullOrWhiteSpace(login)) throw new ArgumentNullException(nameof(login));

            var sharePointUser = _users.FirstOrDefault(u => u.Login.Equals(login, StringComparison.InvariantCulture));
            if (sharePointUser != null) return sharePointUser;

            lock (_lock)
            {
                sharePointUser = _users.FirstOrDefault(u => u.Login.Equals(login, StringComparison.InvariantCulture));
                if (sharePointUser != null) return sharePointUser;

                var user = ClientContext.Web.EnsureUser(login);
                ClientContext.Load(user, u => u.Id, u => u.Email, u => u.Title);
                ClientContext.ExecuteQuery();

                sharePointUser = new SharePointUser(user.Id, login, user.Email, user.Title);

                _users.Add(sharePointUser);

                return sharePointUser;
            }
        }

        protected override internal async Task<int> AddItemAsync<TSharePointItem>(IReadOnlyDictionary<string, object> fields)
        {
            var itemCreateInfo = new ListItemCreationInformation();
            var listItem = ClientContext.GetList(CreateSharePointItem<TSharePointItem>().ListName).AddItem(itemCreateInfo);

            return await Update(listItem, fields);
        }

        protected override internal async Task UpdateItemAsync<TSharePointItem>(int id, IReadOnlyDictionary<string, object> fields)
        {
            await Update(ClientContext.GetList(CreateSharePointItem<TSharePointItem>().ListName).GetItemById(id), fields);
        }

        protected override internal async Task DeleteItemAsync<TSharePointItem>(int id)
        {
            await Delete<TSharePointItem>(id);
        }

        protected override internal async Task<KeyValuePair<ISharePointMetadata, Dictionary<string, object>>> GetAllFieldsByIdAsync<TSharePointMetadata>(int id)
        {
            TSharePointMetadata item = default;
            var fields = new Dictionary<string, object>();

            using (var clientContext = new ClientContext(_webFullUrl))
            {
                clientContext.Credentials = _networkCredential;

                var listName = CreateSharePointItem<TSharePointMetadata>().ListName;
                var listWithContent = clientContext.GetList(listName);

                var listItemCollection = listWithContent.GetItems(new CamlQuery()
                {
                    ViewXml = string.Format(SharePointQueries.QueryItemById, id)
                });
                clientContext.Load(listItemCollection);
                await clientContext.ExecuteQueryAsync();

                if (listItemCollection != null)
                {
                    var items = listItemCollection.ToKnowType<TSharePointMetadata>(clientContext);
                    var listItem = listItemCollection?.FirstOrDefault();

                    item = items.FirstOrDefault();
                    listItem?.FieldValues.Keys.ToList().ForEach(key =>
                    {
                        fields.Add(key, listItem[key]);
                    });
                }
            }

            return new KeyValuePair<ISharePointMetadata, Dictionary<string, object>>(item, fields);
        }

        protected override internal async Task<ICollection<TSharePointMetadata>> GetItemsAsync<TSharePointMetadata>(string viewXml)
        {
            CamlQuery camlQuery = null;
            if (!string.IsNullOrWhiteSpace(viewXml))
            {
                camlQuery = new CamlQuery() { ViewXml = viewXml };
            }

            return await GetItemsByCamlQuery<TSharePointMetadata>(camlQuery);
        }

        protected override internal async Task<(int Id, List<string> CreatedFolders)> AddFileAsync<TSharePointFile>(IReadOnlyDictionary<string, object> fields,
            string fileName, string folderName, Stream content, bool isUpdateFile)
        {
            var fileCreateInfo = new FileCreationInformation
            {
                Url = fileName,
                ContentStream = content,
                Overwrite = isUpdateFile
            };

            var documentLibrary = ClientContext.GetList(CreateSharePointItem<TSharePointFile>().ListName);

            var folder = documentLibrary.RootFolder;

            var baseFolder = "";
            var createdFolders = new List<string>();
            foreach (var name in (folderName ?? string.Empty).Split('/')
                .Where(name => !string.IsNullOrWhiteSpace(name)))
            {
                baseFolder += $"/{name}";

                ClientContext.Load(folder.Folders);
                await ClientContext.ExecuteQueryAsync();

                if (!folder.Folders.Any(f => f.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase)))
                {
                    folder = folder.Folders.Add(name);
                    await ClientContext.ExecuteQueryAsync();

                    createdFolders.Add(baseFolder.Substring(1));
                }
                else
                {
                    folder = folder.Folders.Single(f => f.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
                }
            }

            var file = folder.Files.Add(fileCreateInfo);
            await ClientContext.ExecuteQueryAsync();

            ClientContext.Load(file, f => f.ListItemAllFields[SharePointConstants.FieldNameId]);
            await ClientContext.ExecuteQueryAsync();

            var id = (int)file.ListItemAllFields[SharePointConstants.FieldNameId];
            try
            {
                id = await Update(file.ListItemAllFields, fields);
            } 
            catch (Exception ex)
            {
                throw new SharePointException(SharePointErrorCode.UpdatedDocumentMetadata,
                    new ValueTuple<int, List<string>>(id, createdFolders), ex.Message, ex);
            }

            return (id, createdFolders);
        }

        protected override internal async Task DeleteFileAsync<TSharePointFile>(int id)
        {
            await Delete<TSharePointFile>(id);
        }

        protected override internal async Task RemoveFoldersAsync<TSharePointFile>(List<string> folders)
        {
            if (folders?.Count == 0) return;

            var documentLibrary = ClientContext.GetList(CreateSharePointItem<TSharePointFile>().ListName);

            folders.Reverse();
            foreach (var folderPath in folders)
            {
                var folder = documentLibrary.RootFolder;
                foreach (var name in (folderPath ?? string.Empty).Split('/'))
                {
                    ClientContext.Load(folder.Folders);
                    await ClientContext.ExecuteQueryAsync();

                    if (!folder.Folders.Any(f => f.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))) 
                    {
                        folder = null;
                        break;
                    }
                    folder = folder.Folders.Single(f => f.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
                }
                if (folder != null && !documentLibrary.RootFolder.Equals(folder))
                {
                    folder.DeleteObject();
                }
            }

            await ClientContext.ExecuteQueryAsync();
        }

        public override async Task<SharePointDocumentInfo> GetDocumentsInfoAsync(string documentLibraryName, string fileRef)
        {
            SharePointDocumentInfo documentInfo = null;

            using (var clientContext = new ClientContext(_webFullUrl))
            {
                clientContext.Credentials = _networkCredential;

                var documentLibrary = clientContext.GetList(documentLibraryName);

                var documentRelativeUri = GetDocumentUri(fileRef);
                if (await clientContext.DocumentIsFile(documentLibrary, documentRelativeUri))
                {
                    documentInfo = await clientContext.GetFileInfoByServerRelativeUrl(documentRelativeUri);
                }
                else
                {
                    var folderName = documentRelativeUri.Split('/').Last();
                    var rootPath = documentRelativeUri.Replace(folderName, "");

                    var documents = (await clientContext.GetAllContentByServerRelativeUrl(documentLibrary, documentRelativeUri))
                        .Select(document => new
                        {
                            Id = document.Id,
                            Isfolder = document.FileSystemObjectType == FileSystemObjectType.Folder,
                            FileRef = document.FieldValues[SharePointConstants.FieldNameFileRef].ToString(),
                            Name = document.FieldValues[SharePointConstants.FieldNameFileRef].ToString().Split('/').Last(),
                            Owner = (document.FieldValues[SharePointConstants.FieldNameFileDirRef].ToString()).Replace(rootPath, ""),
                            Level = (document.FieldValues[SharePointConstants.FieldNameFileDirRef].ToString()).Replace(documentRelativeUri, "").Split('/').Length
                        });

                    documentInfo = new SharePointDocumentInfo(0, folderName, null, false);
                    var folders = new Dictionary<string, SharePointDocumentInfo>();
                    folders.Add(folderName, documentInfo);

                    documents.Where(document => document.Isfolder)
                        .OrderBy(document => document.Level)
                        .ToList().ForEach(document =>
                        {
                            var sharePointFileInfo = new SharePointDocumentInfo(document.Id, document.Name, null, false);
                            folders[document.Owner].AddDocument(sharePointFileInfo);

                            folders.Add($"{document.Owner}/{document.Name}", sharePointFileInfo);
                        });

                    foreach (var document in documents.Where(document => !document.Isfolder).ToList())
                    {
                        folders[document.Owner].AddDocument(await clientContext.GetFileInfoByServerRelativeUrl(document.FileRef));
                    }
                }
            }

            return documentInfo;
        }

        #endregion

        #region Methods

        public ICollection<TSharePointFile> GetFiles<TSharePointFile>(CamlQuery camlQuery = null,
            params Expression<Func<ListItemCollection, object>>[] retrievals) where TSharePointFile : ISharePointFile, new()
        {
            return Task<Task<ICollection<TSharePointFile>>>.Factory.StartNew(() => GetFilesAsync<TSharePointFile>(camlQuery, retrievals),
                CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).Unwrap().GetAwaiter().GetResult();
        }

        public async Task<ICollection<TSharePointFile>> GetFilesAsync<TSharePointFile>(CamlQuery camlQuery = null,
            params Expression<Func<ListItemCollection, object>>[] retrievals) where TSharePointFile : ISharePointFile, new()
        {
            return await GetItemsByCamlQuery<TSharePointFile>(camlQuery, retrievals);
        }

        public ICollection<TSharePointItem> GetItems<TSharePointItem>(CamlQuery camlQuery = null,
            params Expression<Func<ListItemCollection, object>>[] retrievals) where TSharePointItem : ISharePointItem, new()
        {
            return Task<Task<ICollection<TSharePointItem>>>.Factory.StartNew(() => GetItemsAsync<TSharePointItem>(camlQuery, retrievals),
                CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).Unwrap().GetAwaiter().GetResult();
        }

        public async Task<ICollection<TSharePointItem>> GetItemsAsync<TSharePointItem>(CamlQuery camlQuery = null,
            params Expression<Func<ListItemCollection, object>>[] retrievals) where TSharePointItem : ISharePointItem, new()
        {
            return await GetItemsByCamlQuery<TSharePointItem>(camlQuery, retrievals);
        }

        protected async Task<ICollection<TSharePointMetadata>> GetItemsByCamlQuery<TSharePointMetadata>(CamlQuery camlQuery = null,
            params Expression<Func<ListItemCollection, object>>[] retrievals) where TSharePointMetadata : ISharePointMetadata, new()
        {
            List<TSharePointMetadata> items = null;

            using (var clientContext = new ClientContext(_webFullUrl))
            {
                clientContext.Credentials = _networkCredential;

                var listName = CreateSharePointItem<TSharePointMetadata>().ListName;
                var listWithContent = clientContext.GetList(listName);

                var listItemCollection = listWithContent.GetItems(camlQuery ?? CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listItemCollection, retrievals);
                await clientContext.ExecuteQueryAsync();

                if (listItemCollection != null)
                {
                    items = listItemCollection.ToKnowType<TSharePointMetadata>(clientContext);
                }
            }

            return items;
        }

        protected Uri GetBaseSharePointUri() => new Uri(new Uri(_webFullUrl).GetLeftPart(UriPartial.Path));

        protected string GetDocumentUri(string documentUri) =>
            Uri.IsWellFormedUriString(documentUri, UriKind.Relative)
                ? documentUri
                : Uri.UnescapeDataString(new Uri(GetBaseSharePointUri(), documentUri).AbsolutePath);

        protected async Task Delete<TSharePointMetadata>(int id) where TSharePointMetadata : ISharePointMetadata, new()
        {
            var listItem = ClientContext.GetList(CreateSharePointItem<TSharePointMetadata>().ListName).GetItemById(id);

            listItem.DeleteObject();
            await ClientContext.ExecuteQueryAsync();
        }

        protected async Task<int> Update(ListItem listItem, IReadOnlyDictionary<string, object> fields)
        {
            foreach (var field in fields.Where(keyValue => !IgnorePropertiesInsertOrUpdate.Contains(keyValue.Key)))
            {
                listItem[field.Key] = field.Value;
            }

            listItem.Update();

            await ClientContext.ExecuteQueryAsync();

            ClientContext.Load(listItem);
            await ClientContext.ExecuteQueryAsync();

            return (int)listItem[SharePointConstants.FieldNameId];
        }

        #endregion

        #region IDisposable - Members

        protected override void Dispose(bool disposing)
        {
            if (Disposed) return;

            if (disposing)
            {
                _clientContext?.Dispose();
            }

            base.Dispose(disposing);
        }

        #endregion
    }
}
