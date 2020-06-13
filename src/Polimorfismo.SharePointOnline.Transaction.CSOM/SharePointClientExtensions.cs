﻿// Copyright 2020 Polimorfismo - José Mauro da Silva Sandy
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
using System.Reflection;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Resources;

namespace Polimorfismo.Microsoft.SharePoint.Transaction
{
    /// <summary>
    /// SharePoint context extension to extend its functionality.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-25 10:04:15 PM</Date>
    internal static class SharePointClientExtensions
    {
        public static TSharePointItem CopyListItemTo<TSharePointItem>(this ClientContext clientContext, ListItem listItem) where TSharePointItem : ISharePointItem
        {
            if (listItem == null)
            {
                throw new ArgumentNullException("listItem");
            }

            var item = Activator.CreateInstance<TSharePointItem>();
            item.Id = listItem.Id;

            foreach (var property in item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetCustomAttributes<SharePointFieldAttribute>().Any(a => !a.IsReference))
                .ToList())
            {
                var attribute = property.GetCustomAttributes<SharePointFieldAttribute>().First();

                if (listItem.FieldValues.ContainsKey(attribute.Name))
                {
                    var itemValue = listItem.FieldValues[attribute.Name];
                    if (itemValue != null)
                    {
                        if (itemValue is FieldUserValue)
                        {
                            property.SetValue(item, attribute.IsUserValue
                                ? ((FieldUserValue)itemValue).LookupValue
                                : (object)((FieldUserValue)itemValue).LookupId);
                        }
                        else if (itemValue is FieldLookupValue)
                        {
                            property.SetValue(item, attribute.IsLookupValue
                                ? ((FieldLookupValue)itemValue).LookupValue
                                : (object)((FieldLookupValue)itemValue).LookupId);
                        }
                        else if (itemValue is FieldUrlValue)
                        {
                            property.SetValue(item, ((FieldUrlValue)itemValue).Url);
                        }
                        else
                        {
                            if (attribute.Type != null)
                            {
                                itemValue = Convert.ChangeType(itemValue, attribute.Type);
                            }
                            property.SetValue(item, itemValue);
                        }
                    }
                }
            }

            return item;
        }

        public static List<TSharePointItem> ToKnowType<TSharePointItem>(this ListItemCollection listItemCollection, ClientContext clientContext)
            where TSharePointItem : ISharePointItem
        {
            var items = Activator.CreateInstance<List<TSharePointItem>>();

            if (listItemCollection != null)
            {
                foreach (var listItem in listItemCollection)
                {
                    items.Add(CopyListItemTo<TSharePointItem>(clientContext, listItem));
                }
            }

            return items;
        }

        public static List GetList(this ClientContext clientContext, string listName)
        {
            if (string.IsNullOrEmpty(listName)) throw new ArgumentNullException(nameof(listName));

            return clientContext.Web.Lists.GetByTitle(listName);
        }

        public static async Task<bool> DocumentIsFile(this ClientContext clientContext, List documentLibrary, string fileRef)
        {
            var camlQuery = new CamlQuery
            {
                ViewXml = string.Format(SharePointQueries.QueryDocumentType, fileRef)
            };

            var listItemCollection = documentLibrary.GetItems(camlQuery);

            clientContext.Load(listItemCollection);
            await clientContext.ExecuteQueryAsync();

            var item = listItemCollection.FirstOrDefault();
            if (item == null)
            {
                throw new SharePointException(SharePointErrorCode.DocumentNotFound, string.Format(SharePointMessages.ERR401, fileRef));
            }

            return item.FileSystemObjectType == FileSystemObjectType.File;
        }

        public static async Task<byte[]> GetContentFile(this ClientContext clientContext, File file)
        {
            byte[] content = null;
            var data = file.OpenBinaryStream();

            if (data != null)
            {
                clientContext.Load(file);
                await clientContext.ExecuteQueryAsync();

                using (var ms = new System.IO.MemoryStream())
                {
                    await data.Value.CopyToAsync(ms);
                    content = ms.ToArray();
                }
            }

            return content;
        }

        public static async Task<SharePointDocumentInfo> GetFileInfoByServerRelativeUrl(this ClientContext clientContext, string fileRef)
        {
            var file = clientContext.Web.GetFileByServerRelativeUrl(fileRef);
            clientContext.Load(file, f => f.Name, f => f.ListItemAllFields.Id);
            await clientContext.ExecuteQueryAsync();

            return new SharePointDocumentInfo(file.ListItemAllFields.Id, 
                                              file.Name, 
                                              await clientContext.GetContentFile(file), 
                                              true);
        }
    }
}
