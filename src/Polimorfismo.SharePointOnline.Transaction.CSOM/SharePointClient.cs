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
using Polimorfismo.SharePoint.Transaction.Logging;

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

        public SharePointClient(string webFullUrl, 
                                ICredentials credential, 
                                ISharePointLogging logging = null)
            : base(logging)
        {
            _webFullUrl = webFullUrl;
            _networkCredential = credential;
        }

        ~SharePointClient() => Dispose(false);

        #endregion

        #region SharePointClientBase - Members

        public override async Task<SharePointUser> GetUserByLogin(string login)
        {
            Logger?.Info($"{GetType().Name}|GetUserByLogin(login = \"{login}\")|Begin");

            if (string.IsNullOrWhiteSpace(login)) throw new ArgumentNullException(nameof(login));

            var sharePointUser = _users.FirstOrDefault(u => u.Login.Equals(login, StringComparison.InvariantCulture));
            if (sharePointUser != null) return sharePointUser;

            var user = ClientContext.Web.EnsureUser(login);
            ClientContext.Load(user);
            await ClientContext.ExecuteQueryAsync();

            sharePointUser = new SharePointUser(user.Id, login, user.Email, user.Title);

            _users.Add(sharePointUser);

            Logger?.Info($"{GetType().Name}|GetUserByLogin(login = \"{login}\")|End");

            return sharePointUser;
        }

        protected override async Task<ICollection<TSharePointItem>> GetItems<TSharePointItem>(string viewXml)
        {
            Logger?.Info($"{GetType().Name}|GetItems(viewXml = \"{viewXml}\")|Begin");

            CamlQuery camlQuery = null;
            if (!string.IsNullOrWhiteSpace(viewXml))
            {
                camlQuery = new CamlQuery() { ViewXml = viewXml };
            }

            Logger?.Info($"{GetType().Name}|GetItems(viewXml = \"{viewXml}\")|End");

            return await GetItems<TSharePointItem>(camlQuery);
        }

        protected override async Task DeleteItem<TSharePointItem>(int id)
        {
            Logger?.Info($"{GetType().Name}|DeleteItem(id = {id})|Begin");

            var listItem = GetList(CreateSharePointItem<TSharePointItem>().ListName).GetItemById(id);
            listItem.DeleteObject();

            await ClientContext.ExecuteQueryAsync();

            Logger?.Info($"{GetType().Name}|DeleteItem(id = {id})|End");
        }

        protected override async Task UpdateItem<TSharePointItem>(int id, IReadOnlyDictionary<string, object> fields)
        {
            Logger?.Info($"{GetType().Name}|UpdateItem( d = {id})|Begin");

            var listName = CreateSharePointItem<TSharePointItem>().ListName;
            await Update(listName, GetList(listName).GetItemById(id), fields);

            Logger?.Info($"{GetType().Name}|UpdateItem(id = {id})|End");
        }

        protected override async Task<int> InsertItem<TSharePointItem>(IReadOnlyDictionary<string, object> fields)
        {
            Logger?.Info($"{GetType().Name}|InsertItem()|Begin");

            var itemCreateInfo = new ListItemCreationInformation();
            var listName = CreateSharePointItem<TSharePointItem>().ListName;
            var listItem = GetList(listName).AddItem(itemCreateInfo);

            var id = await Update(listName, listItem, fields);

            Logger?.Info($"{GetType().Name}|InsertItem()|End|{id}");

            return id;
        }

        #endregion

        #region Methods

        public async Task<ICollection<TSharePointItem>> GetItems<TSharePointItem>(CamlQuery camlQuery = null,
            params Expression<Func<ListItemCollection, object>>[] retrievals) where TSharePointItem : ISharePointItem, new()
        {
            Logger?.Info($"{GetType().Name}|GetItems(camlQuery\"{camlQuery}\")|Begin");

            List<TSharePointItem> items = null;
            ListItemCollection listItemCollection;

            using (var clientContext = new ClientContext(_webFullUrl))
            {
                clientContext.Credentials = _networkCredential;

                var listWithContent = clientContext.Web.Lists.GetByTitle(CreateSharePointItem<TSharePointItem>().ListName);
                clientContext.Load(listWithContent);
                await clientContext.ExecuteQueryAsync();

                listItemCollection = listWithContent.GetItems(camlQuery ?? CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listItemCollection, retrievals);
                await clientContext.ExecuteQueryAsync();

                if (listItemCollection != null)
                {
                    items = listItemCollection.ToKnowType<TSharePointItem>(clientContext);
                }
            }

            Logger?.Info($"{GetType().Name}|GetItems(camlQuery\"{camlQuery}\")|End");

            return items;
        }

        protected List GetList(string listName)
        {
            if (string.IsNullOrEmpty(listName)) throw new ArgumentNullException(nameof(listName));

            return ClientContext.Web.Lists.GetByTitle(listName);
        }

        protected async Task<int> Update(string listName, ListItem listItem, IReadOnlyDictionary<string, object> fields)
        {
            foreach (var field in fields.Where(keyValue => !IgnorePropertiesInsertOrUpdate.Contains(keyValue.Key)))
            {
                listItem[field.Key] = field.Value;

                Logger?.Debug($"{GetType().Name}|Update Field|{listName} -> {field.Key} = {field.Value}");
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
