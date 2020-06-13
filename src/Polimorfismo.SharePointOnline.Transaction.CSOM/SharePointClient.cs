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
using System.Net;
using System.Linq;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.Microsoft.SharePoint.Transaction
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

        public async Task<SharePointDocumentInfo> GetDocuments(string documentLibraryName, string fileRef)
        {
            SharePointDocumentInfo documentInfo = null;

            using (var clientContext = new ClientContext(_webFullUrl))
            {
                clientContext.Credentials = _networkCredential;

                var documentLibrary = clientContext.GetList(documentLibraryName);

                var fileRelativeUri = GetDocumentUri(fileRef);
                if (await clientContext.DocumentIsFile(documentLibrary, fileRelativeUri))
                {
                    documentInfo = await clientContext.GetFileInfoByServerRelativeUrl(fileRelativeUri);
                }
            }

            return documentInfo;
        }

        public override async Task<SharePointUser> GetUserByLogin(string login)
        {
            if (string.IsNullOrWhiteSpace(login)) throw new ArgumentNullException(nameof(login));

            var sharePointUser = _users.FirstOrDefault(u => u.Login.Equals(login, StringComparison.InvariantCulture));
            if (sharePointUser != null) return sharePointUser;

            var user = ClientContext.Web.EnsureUser(login);
            ClientContext.Load(user);
            await ClientContext.ExecuteQueryAsync();

            sharePointUser = new SharePointUser(user.Id, login, user.Email, user.Title);

            _users.Add(sharePointUser);

            return sharePointUser;
        }

        protected override async Task<ICollection<TSharePointItem>> GetItems<TSharePointItem>(string viewXml)
        {
            CamlQuery camlQuery = null;
            if (!string.IsNullOrWhiteSpace(viewXml))
            {
                camlQuery = new CamlQuery() { ViewXml = viewXml };
            }

            return await GetItems<TSharePointItem>(camlQuery);
        }

        protected override async Task DeleteItem<TSharePointItem>(int id)
        {
            var listItem = ClientContext.GetList(CreateSharePointItem<TSharePointItem>().ListName).GetItemById(id);
            listItem.DeleteObject();

            await ClientContext.ExecuteQueryAsync();
        }

        protected override async Task UpdateItem<TSharePointItem>(int id, IReadOnlyDictionary<string, object> fields)
        {
            var listName = CreateSharePointItem<TSharePointItem>().ListName;
            await Update(listName, ClientContext.GetList(listName).GetItemById(id), fields);
        }


        protected override async Task<int> InsertItem<TSharePointItem>(IReadOnlyDictionary<string, object> fields)
        {
            var itemCreateInfo = new ListItemCreationInformation();
            var listName = CreateSharePointItem<TSharePointItem>().ListName;
            var listItem = ClientContext.GetList(listName).AddItem(itemCreateInfo);

            var id = await Update(listName, listItem, fields);

            return id;
        }

        #endregion

        #region Methods
        
        public async Task<ICollection<TSharePointItem>> GetItems<TSharePointItem>(CamlQuery camlQuery = null,
            params Expression<Func<ListItemCollection, object>>[] retrievals) where TSharePointItem : ISharePointItem, new()
        {
            List<TSharePointItem> items = null;
            ListItemCollection listItemCollection;

            using (var clientContext = new ClientContext(_webFullUrl))
            {
                clientContext.Credentials = _networkCredential;

                var listWithContent = ClientContext.GetList(CreateSharePointItem<TSharePointItem>().ListName);

                listItemCollection = listWithContent.GetItems(camlQuery ?? CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listItemCollection, retrievals);
                await clientContext.ExecuteQueryAsync();

                if (listItemCollection != null)
                {
                    items = listItemCollection.ToKnowType<TSharePointItem>(clientContext);
                }
            }
            
            return items;
        }
        
        protected string GetDocumentUri(string documentUri) => 
            Uri.IsWellFormedUriString(documentUri, UriKind.Relative) 
                ? documentUri 
                : Uri.UnescapeDataString(new Uri(GetBaseSharePointUri(), documentUri).AbsolutePath);

        protected Uri GetBaseSharePointUri() => new Uri(new Uri(_webFullUrl).GetLeftPart(UriPartial.Path));
        protected async Task<int> Update(string listName, ListItem listItem, IReadOnlyDictionary<string, object> fields)
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

        public override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            if (disposing)
            {
                _clientContext?.Dispose();
            }
        }

        #endregion
    }
}
