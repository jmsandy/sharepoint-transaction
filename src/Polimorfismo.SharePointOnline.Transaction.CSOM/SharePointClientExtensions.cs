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
using System.Reflection;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using Polimorfismo.SharePoint.Transaction;

namespace Polimorfismo.Microsoft.SharePoint.Transaction
{
    /// <summary>
    /// SharePoint context extension to extend its functionality.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-25 10:04:15 PM</Date>
    internal static class SharePointClientExtensions
    {
        public static User GetUser(this ClientContext clientContext, string logonName)
        {
            var user = clientContext.Web.EnsureUser(logonName);
            clientContext.Load(user);
            clientContext.ExecuteQuery();

            return user;
        }

        public static TSharePointItem CopyListItemTo<TSharePointItem>(this ClientContext clientContext, ListItem listItem) where TSharePointItem : ISharePointItem
        {
            if (listItem == null)
            {
                throw new ArgumentNullException("listItem");
            }

            var item = Activator.CreateInstance<TSharePointItem>();
            foreach (var property in item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetCustomAttributes<SharePointFieldAttribute>().Any()).ToList())
            {
                var attribute = property.GetCustomAttributes<SharePointFieldAttribute>().First();

                if (listItem.FieldValues.ContainsKey(attribute.Name))
                {
                    var itemValue = listItem.FieldValues[attribute.Name];
                    if (itemValue != null)
                    {
                        if (itemValue is FieldLookupValue)
                        {
                            property.SetValue(item, attribute.IsLookupValue
                                ? ((FieldLookupValue)itemValue).LookupValue
                                : (object)((FieldLookupValue)itemValue).LookupId);
                        }
                        else if (itemValue is FieldUrlValue)
                        {
                            property.SetValue(item, ((FieldUrlValue)itemValue).Url);
                        }
                        else if (itemValue is FieldUserValue)
                        {
                            property.SetValue(item, attribute.IsLookupValue
                                ? ((FieldUserValue)itemValue).LookupValue
                                : (object)((FieldUserValue)itemValue).LookupId);
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
    }
}
