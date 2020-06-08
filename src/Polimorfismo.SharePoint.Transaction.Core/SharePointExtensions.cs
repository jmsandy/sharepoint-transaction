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

using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Extension class to assist operations performed in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 09:36:16 PM</Date>
    internal static class SharePointExtensions
    {
        #region Extensions - ISharePointItem

        public static Dictionary<string, object> ToDictionary(this ISharePointItem item)
        {
            return SharePointReflectionUtils.GetSharePointDictionaryValues(item);
        }

        public static Dictionary<string, object> GetReferences(this ISharePointItem item)
        {
            return SharePointReflectionUtils.GetSharePointRefencesDictionaryValues(item);
        }

        public static IReadOnlyDictionary<string, object> ConfigureReferences(this SharePointItemTracking itemTracking, SharePointListItemTracking listTracking)
        {
            var fields = itemTracking.Fields.ToDictionary();

            foreach (var reference in itemTracking.Item.GetReferences())
            {
                if (fields.ContainsKey(reference.Key))
                {
                    var value = listTracking.Get(reference.Value as ISharePointItem);
                    if (value != null)
                    {
                        itemTracking.Fields[reference.Key] = value.Id;
                    }
                }
            }

            return fields;
        }

        #endregion
    }
}
