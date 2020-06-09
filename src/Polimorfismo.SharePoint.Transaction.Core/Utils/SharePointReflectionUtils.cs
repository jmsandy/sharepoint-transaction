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

using System.Linq;
using System.Reflection;
using System.Collections.Generic;

namespace Polimorfismo.SharePoint.Transaction.Utils
{
    /// <summary>
    /// Extension to manipulate library metadata.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 09:07:56 PM</Date>
    internal static class SharePointReflectionUtils
    {
        #region Methods

        public static Dictionary<string, object> GetSharePointDictionaryValues<TSharePointItem>(TSharePointItem item) where TSharePointItem : ISharePointItem
        {
            var dictionary = new Dictionary<string, object>();

            foreach (var property in item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetCustomAttributes<SharePointFieldAttribute>().Any(a => !a.IsIgnoreToInsertOrUpdate)).ToList())
            {
                dictionary.Add(property.GetCustomAttributes<SharePointFieldAttribute>().First().Name, property.GetValue(item));
            }

            return dictionary;
        }

        public static Dictionary<string, object> GetSharePointReferencesDictionaryValues<TSharePointItem>(TSharePointItem item) where TSharePointItem : ISharePointItem
        {
            var dictionary = new Dictionary<string, object>();

            foreach (var property in item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetCustomAttributes<SharePointFieldAttribute>().Any(a => a.IsReference)).ToList())
            {
                dictionary.Add(property.GetCustomAttributes<SharePointFieldAttribute>().First().Name, property.GetValue(item));
            }

            return dictionary;
        }

        public static Dictionary<string, object> GetSharePointUsersDictionaryValues<TSharePointItem>(TSharePointItem item) where TSharePointItem : ISharePointItem
        {
            var dictionary = new Dictionary<string, object>();

            foreach (var property in item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetCustomAttributes<SharePointFieldAttribute>().Any(a => a.IsUserValue)).ToList())
            {
                dictionary.Add(property.GetCustomAttributes<SharePointFieldAttribute>().First().Name, property.GetValue(item));
            }

            return dictionary;
        }

        #endregion
    }
}
